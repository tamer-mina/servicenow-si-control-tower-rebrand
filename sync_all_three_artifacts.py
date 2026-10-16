import csv
import html
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

PREV_JSON = "/Users/tamer.mina/Documents/Antigravity/full_sitemap_crawl_results.json"
CHECKPOINT_JSONL = "/Users/tamer.mina/Documents/Antigravity/hybrid_crawl_checkpoint.jsonl"
FINAL_JSON = "/Users/tamer.mina/Documents/Antigravity/hybrid_full_crawl_results.json"
NEW_PDFS_JSON = "/Users/tamer.mina/Documents/Antigravity/new_pdfs_audit.json"
ANALYTICS_XLSX = "/Users/tamer.mina/Downloads/Brand 3.0 page prioritization.xlsx"
PAGE_INSTANCES_PATH = "/Users/tamer.mina/Documents/Antigravity/page_instances_388.json"
REGIONAL_INVENTORY_PATH = "/Users/tamer.mina/Documents/Antigravity/regional_pages_inventory.json"
MASTER_IMAGE_OCR_PATH = "/Users/tamer.mina/Documents/Antigravity/master_image_ocr_inventory.json"
OUT_CSV = "/Users/tamer.mina/Documents/Antigravity/aict-rebrand-inventory.csv"
OUT_HTML = "/Users/tamer.mina/Documents/Antigravity/aict-rebrand-interactive-report.html"

LOCALE_BADGE_MAP = {
    "us": ("US", "b-loc-us"),
    "uk": ("UK", "b-loc"),
    "au": ("AU", "b-loc"),
    "in": ("IN", "b-loc"),
    "de": ("DE", "b-loc"),
    "fr": ("FR", "b-loc"),
    "fr-ca": ("FR-CA", "b-loc"),
    "es": ("ES", "b-loc"),
    "latam": ("LATAM", "b-loc"),
    "br": ("BR", "b-loc"),
    "it": ("IT", "b-loc"),
    "nl": ("NL", "b-loc"),
    "jp": ("JP", "b-loc"),
    "kr": ("KR", "b-loc"),
}

def get_page_locale(url):
    m = re.match(r"^https?://www\.servicenow\.com/(uk|au|in|de|fr-ca|fr|es|latam|br|it|nl|jp|kr)(?=/|$)", url.lower())
    return m.group(1) if m else "us"

def normalize_url(u):
    if u.startswith("/"):
        u = "https://www.servicenow.com" + u
    u = u.split("#")[0].split("?")[0].rstrip()
    if u != "https://www.servicenow.com/" and u.endswith("/"):
        path = u.replace("https://www.servicenow.com", "")
        if len(path.strip("/").split("/")) > 1:
            u = u.rstrip("/")
    return u

def full_url(path_or_url):
    p = path_or_url.strip().rstrip(",)")
    if p.startswith("http://") or p.startswith("https://"):
        return p
    if not p.startswith("/"):
        p = "/" + p
    return "https://www.servicenow.com" + p

def a_link(path_or_url, label=None, color="var(--blue)"):
    href = full_url(path_or_url)
    disp = label if label is not None else href
    return (
        f'<a class="ext-link" href="{html.escape(href)}" target="_blank" rel="noopener noreferrer" '
        f'title="Open {html.escape(href)} in browser">'
        f'{html.escape(disp)}</a>'
    )

# ---------------------------------------------------------------------------
# 1. Parse Analytics & Page Category Workbook (Prioritization + Category Summary)
# ---------------------------------------------------------------------------
def to_float(v):
    try:
        return float(v)
    except Exception:
        return 0.0

def to_int(v):
    try:
        return int(float(v))
    except Exception:
        return 0

analytics_url_map = {}
analytics_url_map_no_html = {}
cat_baseline_summary = {}

if os.path.exists(ANALYTICS_XLSX):
    with zipfile.ZipFile(ANALYTICS_XLSX, "r") as z:
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sst_root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        shared_strings = ["".join(t.text or "" for t in si.findall(".//main:t", ns)) for si in sst_root.findall("main:si", ns)]

        def parse_sheet(sheet_file):
            root = ET.fromstring(z.read(sheet_file))
            rows = []
            for r in root.findall(".//main:sheetData/main:row", ns):
                r_idx = int(r.attrib.get("r", 0))
                row_dict = {}
                for c in r.findall("main:c", ns):
                    ref = c.attrib.get("r", "")
                    col = "".join(ch for ch in ref if ch.isalpha())
                    t = c.attrib.get("t", "")
                    v_el = c.find("main:v", ns)
                    val = v_el.text if v_el is not None else ""
                    if t == "s" and val != "":
                        val = shared_strings[int(val)]
                    row_dict[col] = val
                rows.append((r_idx, row_dict))
            return rows

        cat_rows = parse_sheet("xl/worksheets/sheet7.xml")
        for r_idx, d in cat_rows:
            if r_idx < 2:
                continue
            cname = d.get("B", "").strip()
            if not cname or cname == "Total":
                continue
            cat_baseline_summary[cname] = {
                "strategic_visibility": to_int(d.get("A", 0)),
                "site_urls": to_int(d.get("C", 0)),
                "site_visits": to_int(d.get("D", 0)),
                "site_traffic_share": to_float(d.get("E", 0)),
                "site_conversions": to_int(d.get("F", 0)),
                "site_fcr": to_float(d.get("G", 0)),
            }

        prio_rows = parse_sheet("xl/worksheets/sheet3.xml")
        for r_idx, d in prio_rows:
            if r_idx < 7:
                continue
            u = d.get("B", "").strip()
            if not u:
                continue
            norm = u.lower().replace("https://", "").replace("http://", "").rstrip("/")
            entry = {
                "category": d.get("A", "").strip(),
                "raw_url": u,
                "page_status": d.get("C", "").strip(),
                "ytd_visits": to_int(d.get("D", 0)),
                "ytd_conversions": to_int(d.get("E", 0)),
                "ytd_fcr": to_float(d.get("F", 0)),
                "ytd_rank": to_int(d.get("G", 0)),
                "strategic_visibility": to_int(d.get("J", 0)),
                "traffic_score": to_int(d.get("K", 0)),
                "connected_exp": d.get("L", "").strip(),
                "urgency": to_int(d.get("M", 0)),
                "loe_dev": d.get("N", "").strip(),
                "loe_assets": d.get("O", "").strip(),
                "prio_score": round(to_float(d.get("P", 0)), 2),
            }
            if norm not in analytics_url_map or entry["ytd_visits"] > analytics_url_map[norm]["ytd_visits"]:
                analytics_url_map[norm] = entry
            norm_no_html = re.sub(r"\.html$", "", norm)
            if norm_no_html not in analytics_url_map_no_html or entry["ytd_visits"] > analytics_url_map_no_html[norm_no_html]["ytd_visits"]:
                analytics_url_map_no_html[norm_no_html] = entry

LOCALE_PREFIX_RE = re.compile(
    r"^www\.servicenow\.com/(?:uk|fr|de|es|it|nl|jp|kr|br|pt|latam|apj|au|ca|fr-ca|zh|cn|in|sg|mx|se|ch|at|be|dk|fi|no|pl|ae|sa|za|il|tr|hk|tw)(?=/|$)"
)

def match_page_analytics(url):
    norm = url.lower().replace("https://", "").replace("http://", "").rstrip("/")
    norm_no_html = re.sub(r"\.html$", "", norm)
    if norm in analytics_url_map:
        return {**analytics_url_map[norm], "match_type": "Exact URL"}
    if norm_no_html in analytics_url_map_no_html:
        return {**analytics_url_map_no_html[norm_no_html], "match_type": "Exact URL"}
    en_norm = LOCALE_PREFIX_RE.sub("www.servicenow.com", norm)
    en_norm_no_html = re.sub(r"\.html$", "", en_norm)
    if en_norm in analytics_url_map:
        m = analytics_url_map[en_norm]
        return {
            **m,
            "ytd_visits": 0,
            "ytd_conversions": 0,
            "ytd_fcr": 0.0,
            "ytd_rank": 0,
            "en_baseline_visits": m["ytd_visits"],
            "en_baseline_conversions": m["ytd_conversions"],
            "en_baseline_prio_score": m["prio_score"],
            "match_type": "Regional (EN Canonical Category)"
        }
    if en_norm_no_html in analytics_url_map_no_html:
        m = analytics_url_map_no_html[en_norm_no_html]
        return {
            **m,
            "ytd_visits": 0,
            "ytd_conversions": 0,
            "ytd_fcr": 0.0,
            "ytd_rank": 0,
            "en_baseline_visits": m["ytd_visits"],
            "en_baseline_conversions": m["ytd_conversions"],
            "en_baseline_prio_score": m["prio_score"],
            "match_type": "Regional (EN Canonical Category)"
        }
    if en_norm.endswith("/pricing.html"):
        parent_norm = en_norm.replace("/pricing.html", ".html")
        if parent_norm in analytics_url_map:
            pm = analytics_url_map[parent_norm]
            return {
                "category": pm["category"],
                "raw_url": url,
                "page_status": "Live",
                "ytd_visits": 0,
                "ytd_conversions": 0,
                "ytd_fcr": 0.0,
                "ytd_rank": 0,
                "strategic_visibility": pm["strategic_visibility"],
                "traffic_score": 1,
                "connected_exp": pm["connected_exp"],
                "urgency": pm["urgency"],
                "loe_dev": "Low",
                "loe_assets": "Low",
                "prio_score": round(0.45 * pm["strategic_visibility"] + 0.35 * 1 + 0.20 * pm["urgency"], 2),
                "match_type": "Parent Product Category"
            }
    if "/standard/resource-center/" in en_norm or "/resources/" in en_norm:
        cat, vis, urg = "Ungated ePDF Pages", 3, 1
    elif "/docs/" in en_norm:
        cat, vis, urg = "Documentation, Release Notes & API", 1, 1
    elif "/blogs/" in en_norm:
        cat, vis, urg = "Blogs Pages", 2, 1
    elif "/events/world-forum/" in en_norm:
        cat, vis, urg = "Events - World Forum - City Landing Page", 3, 3
    elif "/events/" in en_norm:
        cat, vis, urg = "Events - Section Pages", 3, 1
    elif "/partners/" in en_norm:
        cat, vis, urg = "Partners", 3, 1
    elif "/customers/" in en_norm:
        cat, vis, urg = "Customer Pages", 4, 1
    elif "/workflow/" in en_norm:
        cat, vis, urg = "Workflow - Other Pages", 4, 3
    elif "/university/" in en_norm:
        cat, vis, urg = "Learning - ServiceNow University", 4, 1
    elif "/products/" in en_norm:
        cat, vis, urg = "Application Pages", 2, 2
    else:
        cat, vis, urg = "Unspecified", 1, 1
    return {
        "category": cat,
        "raw_url": url,
        "page_status": "Live",
        "ytd_visits": 0,
        "ytd_conversions": 0,
        "ytd_fcr": 0.0,
        "ytd_rank": 0,
        "strategic_visibility": vis,
        "traffic_score": 1,
        "connected_exp": "",
        "urgency": urg,
        "loe_dev": "Low",
        "loe_assets": "Low",
        "prio_score": round(0.45 * vis + 0.35 * 1 + 0.20 * urg, 2),
        "match_type": "URL Taxonomy Rule"
    }

def assign_wave(url, m):
    cat = m["category"]
    visits = max(m["ytd_visits"], m.get("en_baseline_visits", 0))
    conv = max(m["ytd_conversions"], m.get("en_baseline_conversions", 0))
    score = max(m["prio_score"], m.get("en_baseline_prio_score", 0.0))
    if (
        "ai-control-tower" in url
        or "ebk-aict-discovery-playbook" in url
        or "contact-us-aict" in url
        or cat in (
            "Homepage", "Platform - Main Page", "Products by Category", "What is ServiceNow",
            "Core AI Pages", "Events - Homepage", "Events - AI Summits Hub Page",
            "Customer - Homepage", "Impact Pages"
        )
        or visits >= 10000
        or conv >= 250
        or score >= 3.33
    ):
        if cat in ("Blogs Pages", "Documentation, Release Notes & API") and visits < 10000:
            pass
        else:
            return "Wave 1", "P0 - Wave 1 (Flagship, Hubs & High-Traffic)", "w1", 0

    if (
        cat in (
            "Product Pages", "Solutions", "Capability Pages", "Application Pages",
            "Platform - Other Pages", "Industry Pages", "What is Pages", "Resource Center",
            "Ungated ePDF Pages", "Events - World Forum - City Landing Page",
            "Events - Section Pages", "Events - Knowledge Other AEM Pages",
            "Contact Us Page", "Company - Executive Team Pages", "Company - All other pages"
        )
        or (cat in ("Workflow - Other Pages", "Workflow - People & Culture") and (visits >= 100 or conv >= 1))
    ):
        return "Wave 2", "P1 - Wave 2 (Commercial & Active GTM)", "w2", 1

    return "Wave 3", "P2 - Wave 3 (Long-Tail Editorial & Archive)", "w3", 2

def map_filter_group(cat):
    if cat in ("Product Pages", "Products by Category"):
        return "cat-product"
    if cat in ("Solutions", "Capability Pages", "Application Pages", "Industry Pages"):
        return "cat-solutions"
    if cat in ("Homepage", "Platform - Main Page", "Platform - Other Pages", "Core AI Pages", "What is ServiceNow", "What is Pages", "Company - Executive Team Pages", "Company - All other pages", "Impact Pages", "Contact Us Page"):
        return "cat-platform"
    if "Events" in cat:
        return "cat-events"
    if cat in ("Resource Center", "Ungated ePDF Pages"):
        return "cat-resources"
    if cat in ("Workflow - Other Pages", "Workflow - People & Culture", "Blogs Pages"):
        return "cat-editorial"
    return "cat-customers"

# ---------------------------------------------------------------------------
# 2. Product Marketing Naming Standard:
#    - Default Everywhere: "Super Intelligence Control Tower" (32 chars, Title Case "Control Tower")
#    - CTA Buttons Only: "Control Tower" (13 chars, e.g. "See Control Tower", "Explore Control Tower")
#    - Skip WWSG body copy character-limit trims; ONLY flag constrained UI headings / <title> tags
#      where "Super Intelligence Control Tower" is way too long so humans can manually decide replacement.
# ---------------------------------------------------------------------------

def make_button_target(cur_text, full_target_text):
    """
    Produces the Product Marketing approved CTA button target using 'Control Tower'
    (e.g., 'See AI Control Tower' -> 'See Control Tower', 'Explore AI Control Tower' -> 'Explore Control Tower',
    'Extend AI control tower' -> 'Extend Control Tower'), plus localized button equivalents.
    """
    full_t = re.sub(r"(?i)\bsuper\s+intelligence\s+control\s+tower\b", "Super Intelligence Control Tower", full_target_text)
    btn_t = re.sub(r"(?i)\b(?:ServiceNow\s+)?Super\s+Intelligence\s+Control\s+Tower\b", "Control Tower", full_t)
    btn_t = re.sub(r"(?i)\b(?:ServiceNow\s+)?AI\s+Control\s+Tower\b", "Control Tower", btn_t)
    localized_btn_swaps = [
        ("Super-Intelligenz-Kontrollturm", "Control Tower"),
        ("Superintelligenz-Kontrollturm", "Control Tower"),
        ("KI-Kontrollturm", "Control Tower"),
        ("Tour de contrôle de la Super Intelligence", "Tour de contrôle"),
        ("Tour de contrôle de la superintelligence", "Tour de contrôle"),
        ("Tour de contrôle Super Intelligence", "Tour de contrôle"),
        ("Tour de contrôle superintelligence", "Tour de contrôle"),
        ("tour de contrôle de la Super Intelligence", "tour de contrôle"),
        ("tour de contrôle de la superintelligence", "tour de contrôle"),
        ("tour de contrôle Super Intelligence", "tour de contrôle"),
        ("tour de contrôle superintelligence", "tour de contrôle"),
        ("Torre de control de Superinteligencia", "Torre de control"),
        ("Torre de control de superinteligencia", "Torre de control"),
        ("torre de control de Superinteligencia", "torre de control"),
        ("torre de control de superinteligencia", "torre de control"),
        ("Torre de controle de Superinteligência", "Torre de controle"),
        ("Torre de controle de superinteligência", "Torre de controle"),
        ("torre de controle de Superinteligência", "torre de controle"),
        ("torre de controle de superinteligência", "torre de controle"),
        ("Torre di controllo della Super Intelligenza", "Torre di controllo"),
        ("Torre di controllo della superintelligenza", "Torre di controllo"),
        ("torre di controllo della Super Intelligenza", "torre di controllo"),
        ("torre di controllo della superintelligenza", "torre di controllo"),
        ("Scopri l'Super Intelligence Control Tower", "Scopri Control Tower"),
        ("Scopri Super Intelligence Control Tower", "Scopri Control Tower"),
        ("Superintelligentie-controletoren", "Controletoren"),
        ("superintelligentie-controletoren", "controletoren"),
        ("controletoren voor Superintelligentie", "controletoren"),
        ("controletoren voor superintelligentie", "controletoren"),
        ("スーパーインテリジェンス・コントロールタワー", "コントロールタワー"),
        ("スーパーインテリジェンス コントロールタワー", "コントロールタワー"),
        ("슈퍼 인텔리전스 컨트롤 타워", "컨트롤 타워"),
        ("슈퍼 인텔리전스 컨트롤타워", "컨트롤 타워"),
    ]
    for full_loc, short_loc in localized_btn_swaps:
        btn_t = btn_t.replace(full_loc, short_loc)
    return btn_t

CTA_BUTTON_TEXT_RE = re.compile(
    r"^(?:See|Explore|Extend|Discover|Try|Get|Watch|View|Learn\s+about|Infos\s+zu|Voir|D[ée]couvrir|Ver|Explorar|Veja|Scopri|Bekijk)\s+.*"
    r"|(?:を見る|보기)$",
    re.I
)

TAGLINE_PATTERNS = (
    "business reinvention",
    "government reinvention",
    "geschäftliche neuausrichtung",
    "réinvention",
    "reinvención",
    "reinvenção",
    "reinvenzione",
    "heruitvinding",
    "ビジネス変革",
    "비즈니스 혁신",
)

def is_tagline_instance(cur_text, tgt_text=""):
    combined = ((cur_text or "") + " " + (tgt_text or "")).lower()
    return any(p in combined for p in TAGLINE_PATTERNS)

def is_cta_button_instance(cur_text, loc_str):
    cur_s = (cur_text or "").strip()
    loc_s = (loc_str or "")
    if len(cur_s) > 40:
        return False
    if any(k in loc_s for k in ("Heading", "Question", "Document <title>", "CTA Banner")):
        return False
    if "arc-button" in loc_s.lower():
        return True
    if CTA_BUTTON_TEXT_RE.search(cur_s):
        return True
    return False

def audit_wwsg_instance(inst, page_url=""):
    loc = inst.get("location", "")
    cat = inst.get("category", "")
    tgt = inst.get("target_text", "").strip()
    cur = inst.get("current_text", "").strip()

    # Always enforce Title Case 'Control Tower' in target text, even when live copy had lowercase 'AI control tower'
    tgt = re.sub(r"(?i)\bsuper\s+intelligence\s+control\s+tower\b", "Super Intelligence Control Tower", tgt)
    tgt = re.sub(r"(?i)\bsi\s+control\s+tower\b", "Super Intelligence Control Tower", tgt)
    tgt = re.sub(r"\ban\s+Super\s+Intelligence\s+Control\s+Tower\b", "a Super Intelligence Control Tower", tgt)
    if cur.startswith("The ") and tgt.startswith("the "):
        tgt = "T" + tgt[1:]
    had_lowercase_ct = bool(re.search(r"control\s+tower", cur))

    is_cta_btn = is_cta_button_instance(cur, loc)
    is_tagline = is_tagline_instance(cur, tgt)
    display_loc = loc
    if is_cta_btn:
        if "arc-heading-block" in loc:
            display_loc = loc.replace("Section Intro Copy (<arc-heading-block> > <p>)", "Hero / Section CTA Button (<arc-button> inside <arc-heading-block>)")
        elif "arc-universal-group" in loc:
            display_loc = loc.replace("Benefits Card Copy (<arc-universal-group> > <p>)", "Card CTA Button (<arc-button> inside <arc-universal-group>)")
        elif "arc-feature" in loc:
            display_loc = loc.replace("Interactive Feature Tab Copy (<arc-feature> > <p>)", "Feature Tab CTA Button (<arc-button> inside <arc-feature>)")
        elif "<action-item>" in loc:
            display_loc = loc.replace("Body Copy / Cross-Sell Card (<p> / <action-item>)", "Cross-Sell CTA Button (<arc-button> / <action-item>)")

    # Classify component type and whether it is a constrained heading/title that must be flagged if too long
    limit = None
    comp_name = "Body copy"
    is_constrained_heading_or_title = False
    flag_threshold = None

    if is_cta_btn:
        limit = 25
        comp_name = "CTA button"
    elif is_tagline:
        limit = 175
        comp_name = "Company positioning tagline"
    elif "arc-secondary-nav" in loc:
        limit = 50
        comp_name = "Secondary nav title"
        is_constrained_heading_or_title = True
        flag_threshold = 50
    elif "arc-marquee-large" in loc and ("heading" in loc.lower() or "Eyebrow" in loc):
        limit = 50
        comp_name = "Marquee large heading"
        is_constrained_heading_or_title = True
        flag_threshold = 50
    elif "arc-marquee-large" in loc and "<p>" in loc:
        comp_name = "Marquee large description"
    elif "arc-marquee" in loc and ("heading" in loc.lower() or "H1" in loc):
        limit = 50
        comp_name = "Marquee heading"
        is_constrained_heading_or_title = True
        flag_threshold = 50
    elif "arc-marquee" in loc and "<p>" in loc:
        comp_name = "Marquee description"
    elif "arc-accordion-item" in loc and "heading" in loc:
        limit = 50
        comp_name = "FAQ accordion heading"
        is_constrained_heading_or_title = True
        flag_threshold = 60  # Short 41-45 char FAQ questions fit; flag long 65-115 char FAQ questions
    elif "arc-accordion-item" in loc and "copy-block" in loc:
        comp_name = "Accordion body"
    elif "arc-next-steps" in loc and "<p>" in loc:
        comp_name = "Next steps description"
    elif "action-item" in loc and "<p>" in loc:
        comp_name = "Next steps action item"
    elif "arc-universal-group" in loc and "<p>" in loc:
        comp_name = "Card description"
    elif ("Heading" in loc or "H1" in loc or "H2" in loc or "H3" in loc) and "Body Copy" not in loc:
        limit = 50
        comp_name = "UI / Section heading"
        is_constrained_heading_or_title = True
        flag_threshold = 50
    elif "Breadcrumb" in loc:
        limit = 50
        comp_name = "Breadcrumb item"
        is_constrained_heading_or_title = True
        flag_threshold = 50
    elif cat == "Image Alt / ARIA Attribute" or "alt=" in loc:
        limit = 125
        comp_name = "Image alt text"
    elif "Document <title>" in loc or "og:title" in loc or "twitter:title" in loc:
        limit = 60
        comp_name = "SEO <title>"
        is_constrained_heading_or_title = True
        flag_threshold = 60
    elif "description" in loc and cat == "Head <title> & <meta>":
        limit = 160
        comp_name = "SEO <meta description>"

    btn_short = make_button_target(cur, tgt)
    t_len = len(tgt)
    btn_len = len(btn_short)

    used_btn_short = False
    recommended_auto_target = tgt
    needs_manual_copy = False
    manual_reason = ""
    is_over = False
    flags = []

    if is_cta_btn:
        used_btn_short = True
        recommended_auto_target = btn_short
        char_check = f"CTA button: [{btn_len}/25 chars — Uses 'Control Tower' per PM direction (Full={t_len} chars)]"
        short_note = f"CTA button: {btn_len}/25 chars ('Control Tower' button standard)"
        flags.append(f'[CTA BUTTON: Uses "{btn_short}" ({btn_len}/25 chars) per Product Marketing button rule]')
    elif is_constrained_heading_or_title and flag_threshold and t_len > flag_threshold:
        is_over = True
        needs_manual_copy = True
        over_by = t_len - limit
        manual_reason = f"Flagged: '{comp_name}' is {t_len}/{limit} chars (+{over_by} over) with 'Super Intelligence Control Tower' — manually decide replacement"
        char_check = f"{comp_name}: [{t_len}/{limit} chars — FLAGGED TOO LONG (+{over_by} chars), manually decide replacement]"
        short_note = f"Flagged: {t_len}/{limit} chars (+{over_by} over) — decide replacement (e.g. 'Control Tower' or trim)"
        flags.append(f"[FLAGGED — NAME TOO LONG: {t_len}/{limit} chars in {comp_name}; manually decide replacement]")
    elif limit:
        char_check = f"{comp_name}: [{t_len}/{limit} chars — Full Name Fits]"
        short_note = f"{comp_name}: {t_len}/{limit} chars"
    else:
        char_check = f"{comp_name}: [{t_len} chars — Direct swap to 'Super Intelligence Control Tower']"
        short_note = f"{comp_name}: {t_len} chars"

    if had_lowercase_ct and not is_cta_btn:
        short_note += " • Capitalize 'Control Tower'"
        char_check += " [Title-Case 'Control Tower']"

    if re.search(r"\b(?:AICT|SICT)\b", cur):
        flags.append('[ACRONYM: Replace "AICT" with "Super Intelligence Control Tower"]')

    has_an_article = bool(re.search(r"\ban\s+AI\s+Control\s+Tower\b", cur, flags=re.I))
    if has_an_article:
        flags.append('[GRAMMAR: "an AI Control Tower" -> "a Super Intelligence Control Tower"]')

    rw_len = len(recommended_auto_target)
    rw_badge = f"[{rw_len}/{limit} chars]" if limit else f"[{rw_len} chars]"

    return {
        "char_check": char_check,
        "short_note": short_note,
        "is_over": is_over,
        "used_si_short": used_btn_short,
        "si_short_target": btn_short,
        "recommended_auto_target": recommended_auto_target,
        "flags": flags,
        "wwsg_rewrite": None,
        "wwsg_rewrite_badge": rw_badge,
        "needs_manual_copy": needs_manual_copy,
        "manual_reason": manual_reason,
        "comp_name": comp_name,
        "limit": limit,
        "display_loc": display_loc,
        "is_cta_btn": is_cta_btn,
    }

# ---------------------------------------------------------------------------
# 3. Load Crawl Data, Regional Pages Inventory, Per-Page Instance Database & Master Image OCR Inventory
# ---------------------------------------------------------------------------
PAGE_INSTANCES_MAP = {}
if os.path.exists(PAGE_INSTANCES_PATH):
    with open(PAGE_INSTANCES_PATH, "r", encoding="utf-8") as _f:
        PAGE_INSTANCES_MAP = json.load(_f)

MASTER_IMAGES_MAP = {}
MASTER_PAGE_TO_IMAGES = {}
if os.path.exists(MASTER_IMAGE_OCR_PATH):
    with open(MASTER_IMAGE_OCR_PATH, "r", encoding="utf-8") as _f:
        _m_ocr = json.load(_f)
        for _im_url, _im_info in _m_ocr.get("images", {}).items():
            if _im_url.lower().endswith(".pdf") or "logo-aicte.svg" in _im_url.lower():
                continue
            _full_im = full_url(_im_url)
            _im_info["url"] = _full_im
            _im_info["has_aict_slug"] = bool(
                ("ai-control-tower" in _full_im.lower() or "aict" in _full_im.lower())
                and "logo-aicte" not in _full_im.lower()
            )
            _raw_pages = [full_url(p) for p in _im_info.get("pages", [])]
            _raw_pages.sort(key=lambda p: (0 if get_page_locale(p) == "us" else 1, p))
            _im_info["host_pages"] = _raw_pages
            _im_info["page_count"] = len(_raw_pages)
            _dims_str = str(_im_info.get("dimensions", ""))
            if "x" in _dims_str.lower():
                try:
                    _w, _h = _dims_str.lower().split("x", 1)
                    _im_info["width"] = int(_w.strip())
                    _im_info["height"] = int(_h.strip())
                except Exception:
                    _im_info["width"] = 0
                    _im_info["height"] = 0
            else:
                _im_info["width"] = _im_info.get("width", 0)
                _im_info["height"] = _im_info.get("height", 0)
            MASTER_IMAGES_MAP[_full_im] = _im_info
        for _pg_url, _im_list in _m_ocr.get("page_to_images", {}).items():
            _clean_list = [full_url(u) for u in _im_list if not u.lower().endswith(".pdf") and "logo-aicte.svg" not in u.lower()]
            if _clean_list:
                MASTER_PAGE_TO_IMAGES[full_url(_pg_url)] = _clean_list

def clean_ocr_display(ocr_lines):
    if isinstance(ocr_lines, str):
        ocr_lines = [ocr_lines]
    cleaned = []
    for ln in (ocr_lines or []):
        c = re.sub(r"\bAl\s+Control\s+Tower\b", "AI Control Tower", str(ln))
        c = re.sub(r"\bAl\s+control\s+tower\b", "AI control tower", c)
        cleaned.append(c)
    return " | ".join(cleaned[:3]) if cleaned else "AI Control Tower"

def get_target_image_url(fim):
    """
    Returns the new DAM image copy path so AEM automation can swap old image path -> new image path.
    - If filename has 'ai-control-tower' or 'aict', replaces with 'super-intelligence-control-tower'
    - Else if filename has 'control-tower', replaces with 'super-intelligence-control-tower'
    - Else inserts '-super-intelligence-control-tower' before rendition (.sm/.xs) or file extension
    """
    if re.search(r"(?i)ai-control-tower|aict", fim):
        return re.sub(r"(?i)ai-control-tower|aict", "super-intelligence-control-tower", fim)
    if re.search(r"(?i)control-tower", fim):
        return re.sub(r"(?i)control-tower", "super-intelligence-control-tower", fim)
    return re.sub(
        r"(\.(?:sm|xs|md|lg|xl)|-(?:sm|xs|md|lg|xl))?(\.[a-zA-Z0-9]+)$",
        r"-super-intelligence-control-tower\1\2",
        fim,
    )

total_baked_images_count = sum(1 for v in MASTER_IMAGES_MAP.values() if v.get("has_baked_aict_pixels"))
total_auto_slug_images_count = sum(1 for v in MASTER_IMAGES_MAP.values() if not v.get("has_baked_aict_pixels") and v.get("has_aict_slug"))
total_zero_pixel_images_count = total_auto_slug_images_count
total_impacted_images_count = total_baked_images_count + total_zero_pixel_images_count

crawled_map = {}
if os.path.exists(PREV_JSON):
    prev = json.load(open(PREV_JSON))
    for r in prev.get("results", []):
        u = normalize_url(r["url"])
        r["url"] = u
        crawled_map[u] = r

if os.path.exists(CHECKPOINT_JSONL):
    with open(CHECKPOINT_JSONL, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                try:
                    r = json.loads(line)
                    u = normalize_url(r["url"])
                    r["url"] = u
                    if u in crawled_map:
                        old = crawled_map[u]
                        for k in ("js_matches", "dam_assets", "alt_matches", "snippets"):
                            if len(old.get(k) or []) > len(r.get(k) or []):
                                r[k] = old[k]
                    crawled_map[u] = r
                except Exception:
                    pass

if os.path.exists(FINAL_JSON):
    try:
        fin = json.load(open(FINAL_JSON))
        for r in fin.get("results", []):
            u = normalize_url(r["url"])
            r["url"] = u
            if u in crawled_map:
                old = crawled_map[u]
                for k in ("js_matches", "dam_assets", "alt_matches", "snippets"):
                    if len(old.get(k) or []) > len(r.get(k) or []):
                        r[k] = old[k]
            crawled_map[u] = r
    except Exception:
        pass

if os.path.exists(REGIONAL_INVENTORY_PATH):
    try:
        reg_list = json.load(open(REGIONAL_INVENTORY_PATH, "r", encoding="utf-8"))
        for r in reg_list:
            u = normalize_url(r["url"])
            r["url"] = u
            if u in crawled_map:
                old = crawled_map[u]
                for k in ("js_matches", "dam_assets", "alt_matches", "snippets", "aict_imgs", "pdfs"):
                    if len(old.get(k) or []) > len(r.get(k) or []):
                        r[k] = old[k]
                for k in ("total_matches", "raw_ai_ct", "raw_aict", "raw_localized", "raw_slug", "true_body"):
                    if (old.get(k) or 0) > (r.get(k) or 0):
                        r[k] = old[k]
            crawled_map[u] = r
    except Exception:
        pass

all_pages = list(crawled_map.values())

def clean_snippet_text(s, page_title=""):
    s = html.unescape(s)
    s = re.sub(r'^.*?-->\s*Skip to Main Content\s*(?:MyNow\s+Products\s+Industries\s+Learning\s+Support\s+Partners\s+Company\s*)?', '', s, flags=re.I)
    s = re.sub(r'^.*?Sales\s+\+1\s*\(844\)\s*863-1987\s*View a Demo\s*', '', s, flags=re.I)
    s = re.sub(r'^.*?Get the latest styles today\.\s*', '', s, flags=re.I)
    s = s.strip()
    if page_title and s.lower() == page_title.lower().strip():
        return ""
    return s

for r in all_pages:
    u = r["url"]
    fu = full_url(u)
    # Merge all discovered images (from Brand 3.0 Excel + component proximity + Vision OCR) for this page
    existing_imgs = [full_url(im.rstrip(",)")) for im in (r.get("aict_imgs") or []) if im.rstrip(",)").endswith((".png", ".jpg", ".jpeg", ".webp", ".svg"))]
    ocr_imgs = MASTER_PAGE_TO_IMAGES.get(fu, [])
    merged_imgs = sorted(set(existing_imgs + ocr_imgs))
    r["aict_imgs"] = merged_imgs

    p_inst = PAGE_INSTANCES_MAP.get(fu, {})
    tm = r.get("total_matches", 0)
    loc = r.get("raw_localized", 0)
    raw_snips = r.get("snippets", [])
    title_clean = (r.get("title") or p_inst.get("title_extracted") or "").replace("--> Skip to Main Content", "").strip()
    r["title_clean"] = title_clean
    r["locale"] = get_page_locale(fu)

    non_nav_snips = []
    for s in raw_snips:
        if (
            "Connect strategy, governance, management, and performance for all you" in s
            or "How does AI Control Tower work" in s
            or '"topic": "ai control tower"' in s
        ):
            continue
        cs = clean_snippet_text(s, title_clean)
        if cs and cs not in non_nav_snips:
            non_nav_snips.append(cs)

    r["non_nav_snips"] = non_nav_snips
    r["has_reinvention_tagline"] = any("business reinvention" in s.lower() for s in non_nav_snips)

    is_legacy_18_nav = (tm == 18 and not non_nav_snips and loc == 0 and "ai-control-tower" not in u.lower() and "aict" not in u.lower() and fu not in PAGE_INSTANCES_MAP)
    is_standard_14_nav = (tm == 14 and not non_nav_snips and loc == 0 and "ai-control-tower" not in u.lower() and "aict" not in u.lower() and fu not in PAGE_INSTANCES_MAP)
    is_12_nav = (tm == 12 and not non_nav_snips and loc == 0 and "ai-control-tower" not in u.lower() and "aict" not in u.lower() and fu not in PAGE_INSTANCES_MAP)

    if is_legacy_18_nav:
        r["chrome_est"] = 18
        r["true_body"] = 0
    elif is_standard_14_nav:
        r["chrome_est"] = 14
        r["true_body"] = 0
    elif is_12_nav:
        r["chrome_est"] = 12
        r["true_body"] = 0
    elif "products/ai-control-tower.html" in u or "events/live/ai-control-tower" in u:
        r["chrome_est"] = 12 if tm >= 12 else 0
        r["true_body"] = max(len(p_inst.get("body_text_instances", [])), tm - r["chrome_est"])
    elif tm > 18 and not non_nav_snips:
        r["chrome_est"] = 18
        r["true_body"] = max(len(p_inst.get("body_text_instances", [])), tm - 18)
    elif tm > 14:
        if any(p in u for p in ("/workflow/", "/blogs/", "/events/", "/customers/", "/standard/resource-center/")):
            r["chrome_est"] = 18 if tm >= 18 else 14
        else:
            r["chrome_est"] = 14
        r["true_body"] = max(1 if non_nav_snips else 0, len(p_inst.get("body_text_instances", [])), tm - r["chrome_est"])
    else:
        r["chrome_est"] = tm if (not non_nav_snips and fu not in PAGE_INSTANCES_MAP) else r.get("chrome_est", 0)
        r["true_body"] = max(1 if non_nav_snips else 0, len(p_inst.get("body_text_instances", [])), r.get("true_body", 0), tm - r["chrome_est"])

    an = match_page_analytics(u)
    r["an"] = an
    wave_lbl, prio_lbl, wave_code, wave_ord = assign_wave(u, an)
    r["wave_lbl"] = wave_lbl
    r["prio_lbl"] = prio_lbl
    r["wave_code"] = wave_code
    r["wave_ord"] = wave_ord

unique_body_pages = [
    r for r in all_pages
    if (
        r["true_body"] > 0
        or r.get("raw_localized", 0) > 0
        or "ai-control-tower" in r["url"].lower()
        or "aict" in r["url"].lower()
        or full_url(r["url"]) in PAGE_INSTANCES_MAP
    )
    and "control-tower-model.html" not in r["url"].lower()
]

unique_body_pages.sort(
    key=lambda x: (
        x["wave_ord"],
        -(1 if "products/ai-control-tower.html" in x["url"] else 0),
        -(1 if x["an"]["category"] == "Homepage" else 0),
        -max(x["an"]["ytd_visits"], x["an"].get("en_baseline_visits", 0)),
        0 if x.get("locale", "us") == "us" else 1,
        x.get("locale", "us"),
        -max(x["an"]["prio_score"], x["an"].get("en_baseline_prio_score", 0.0)),
        -(x["true_body"] + x.get("raw_localized", 0))
    )
)

ssi_only_pages = [r for r in all_pages if r not in unique_body_pages and r.get("total_matches", 0) > 0]

# Map PDFs to host pages and PDF match counts
pdf_to_hosts = {}
for r in all_pages:
    for p in r.get("pdfs", []):
        if "business-continuity-faq" not in p:
            fp = full_url(p)
            pdf_to_hosts.setdefault(fp, []).append(r["url"])

audited_pdfs = []
pdf_match_lookup = {}
if os.path.exists(NEW_PDFS_JSON):
    audited_pdfs = json.load(open(NEW_PDFS_JSON))
    for p in audited_pdfs:
        fp = full_url(p["url"])
        m_cnt = 6 if "ebk-aict-discovery-playbook" in fp else p.get("matches", 0)
        pdf_match_lookup[fp] = (m_cnt, p.get("pages", 0), p.get("meta_title", ""))

legal_pdf_url = "https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/legal/ai-control-tower-servicenow-subscription-unit-overview.pdf"
pdf_match_lookup[legal_pdf_url] = (7, 4, "AI Control Tower - ServiceNow Subscription Unit Overview")

img_map = {}
for r in all_pages:
    for im in r.get("aict_imgs", []):
        clean_im = full_url(im.rstrip(",)"))
        if clean_im.endswith((".png", ".jpg", ".jpeg", ".webp", ".svg")):
            img_map.setdefault(clean_im, []).append((r["url"], r.get("alt_matches", [])))

for _im_url, _im_info in MASTER_IMAGES_MAP.items():
    if _im_url not in img_map:
        img_map[_im_url] = [(pg, []) for pg in _im_info.get("pages", [])]
    else:
        _existing_pgs = {full_url(u) for u, _ in img_map[_im_url]}
        for pg in _im_info.get("pages", []):
            if full_url(pg) not in _existing_pgs:
                img_map[_im_url].append((pg, []))

# Ensure logo-aicte.svg is also in img_map for CSV exclusion note
img_map.setdefault("https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/academic-partnerships/logo-aicte.svg", [("https://www.servicenow.com/company/Filtered-by-location/india/academic-partnerships.html", [])])

# ---------------------------------------------------------------------------
# 4. Split Page Breakdown Helper: Automated Updates (Tab 1) vs. Manual Updates (Tab 2)
# ---------------------------------------------------------------------------
def build_split_page_breakdowns(r):
    """
    Splits all changes for page `r` into Automated vs. Manual tracks, deduplicates repetitive
    <title>/<meta> and <a href> rows, and returns structured 5-column Before -> After mini-table
    rows (`drawer_rows_html`) for the collapsible page detail drawer.
    """
    u = full_url(r["url"])
    loc_code = r.get("locale", get_page_locale(u))
    cat = r.get("an", {}).get("category", "")
    ch = r["chrome_est"]
    raw_ai_ct = r.get("raw_ai_ct", 0)
    raw_aict = r.get("raw_aict", 0)
    raw_loc = r.get("raw_localized", 0)
    raw_slug = r.get("raw_slug", 0)
    body_ai_ct = max(0, raw_ai_ct - (10 if ch >= 12 else 0))
    body_slug = max(0, raw_slug - (4 if ch >= 12 else 0))
    imgs = sorted(set(im.rstrip(",)") for im in (r.get("aict_imgs") or []) if im.rstrip(",)").endswith((".png", ".jpg", ".jpeg", ".webp", ".svg"))))
    pdfs = sorted(set(full_url(p) for p in (r.get("pdfs") or []) if "business-continuity-faq" not in p))

    p_inst = PAGE_INSTANCES_MAP.get(u, {})
    meta_instances = p_inst.get("meta_instances", [])
    body_text_instances = p_inst.get("body_text_instances", [])
    alt_instances = p_inst.get("alt_instances", [])
    data_attr_instances = p_inst.get("data_attr_instances", [])
    body_slug_links = p_inst.get("body_slug_links", [])

    drawer_auto_rows = []
    drawer_manual_rows = []
    auto_csv_items = []
    manual_csv_items = []
    wwsg_csv_notes = []

    auto_text_meta_alt_cnt = 0
    auto_link_cnt = len(body_slug_links)
    manual_copy_cnt = 0
    si_short_cnt = 0
    manual_reasons = []
    manual_filter_tags = set()

    # A. URL Slug / 301 Redirect (Manual / Web Ops task)
    if "ai-control-tower" in u.lower() or "aict" in u.lower():
        new_u = re.sub(r'ai-control-tower', 'super-intelligence-control-tower', u, flags=re.I)
        new_u = re.sub(r'aict', 'super-intelligence-control-tower', new_u, flags=re.I)
        manual_reasons.append("Page Slug Move & 301")
        manual_filter_tags.add("m-slug")
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>Page URL Slug &amp; 301 Redirect</strong><div class='sub-note'>AEM page move + canonical &amp; hreflang</div></td>"
            f"<td class='mono'>{a_link(u)}</td>"
            f"<td class='mono'><code>{html.escape(new_u)}</code></td>"
            f"<td>HTTP 301 Permanent Redirect <span class='link-pill' onclick=\"event.stopPropagation(); switchPrimaryTab('tab-redirects', null, 'btn-tab-redirects')\">View 301 Rules</span></td>"
            f"</tr>"
        )
        manual_csv_items.append(f"[MANUAL: URL SLUG & 301] Move {u} -> {new_u}")

    # B. <title> & <meta> Tag Instances (Automated by default; Flagged Manual if <title> > 60 chars)
    if meta_instances:
        meta_auto_csv_parts = []
        meta_manual_csv_parts = []
        meta_groups = []
        meta_group_map = {}
        for idx, m in enumerate(meta_instances, 1):
            m_loc = m.get("location", "meta")
            m_old = m.get("current_text", "")
            m_new = m.get("target_text", "")
            w_res = audit_wwsg_instance(m, u)
            if w_res["used_si_short"]:
                si_short_cnt += 1
            rec_tgt = w_res["recommended_auto_target"]
            is_man_meta = w_res["needs_manual_copy"]
            if is_man_meta:
                meta_manual_csv_parts.append(f"Meta #{idx} [MANUAL_FLAGGED_TOO_LONG] ({m_loc}) {w_res['char_check']}: \"{m_old}\" -> \"{rec_tgt}\"")
            else:
                auto_text_meta_alt_cnt += 1
                meta_auto_csv_parts.append(f"Meta #{idx} [AUTO] ({m_loc}) {w_res['char_check']}: \"{m_old}\" -> \"{rec_tgt}\"")

            # Clean tag name for grouping
            if "Document <title>" in m_loc:
                short_tag = "<title>"
                grp_kind = "Title"
            elif "og:title" in m_loc:
                short_tag = "og:title"
                grp_kind = "Title"
            elif "twitter:title" in m_loc:
                short_tag = "twitter:title"
                grp_kind = "Title"
            elif "og:description" in m_loc:
                short_tag = "og:description"
                grp_kind = "Description"
            elif "twitter:description" in m_loc:
                short_tag = "twitter:description"
                grp_kind = "Description"
            elif "description" in m_loc:
                short_tag = "meta description"
                grp_kind = "Description"
            else:
                short_tag = m_loc
                grp_kind = "Meta"

            gkey = (grp_kind, m_old, rec_tgt, w_res["short_note"], is_man_meta)
            if gkey not in meta_group_map:
                gobj = {
                    "kind": grp_kind,
                    "tags": [short_tag],
                    "old": m_old,
                    "new": rec_tgt,
                    "note": w_res["short_note"],
                    "manual": is_man_meta,
                }
                meta_group_map[gkey] = gobj
                meta_groups.append(gobj)
            else:
                meta_group_map[gkey]["tags"].append(short_tag)

        for g in meta_groups:
            tag_cnt = len(g["tags"])
            tags_str = ", ".join(html.escape(t) for t in g["tags"])
            lbl = "Page Title Tags" if g["kind"] == "Title" else ("Meta Description Tags" if g["kind"] == "Description" else "Meta Tag")
            cnt_pill = f" ({tag_cnt} tags)" if tag_cnt > 1 else ""
            if g["manual"]:
                manual_copy_cnt += 1
                manual_filter_tags.add("m-copy")
                drawer_manual_rows.append(
                    f"<tr class='detail-item-manual'>"
                    f"<td><span class='badge b-manual'>Manual</span></td>"
                    f"<td><strong>{lbl}{cnt_pill}</strong><div class='sub-note'><code>{tags_str}</code> • Flagged: Name too long — manually decide replacement</div></td>"
                    f"<td>{html.escape(g['old'])}</td>"
                    f"<td><strong>{html.escape(g['new'])}</strong></td>"
                    f"<td>{html.escape(g['note'])}</td>"
                    f"</tr>"
                )
            else:
                drawer_auto_rows.append(
                    f"<tr class='detail-item-auto'>"
                    f"<td><span class='badge b-auto'>Auto</span></td>"
                    f"<td><strong>{lbl}{cnt_pill}</strong><div class='sub-note'><code>{tags_str}</code></div></td>"
                    f"<td>{html.escape(g['old'])}</td>"
                    f"<td><strong>{html.escape(g['new'])}</strong></td>"
                    f"<td>{html.escape(g['note'])}</td>"
                    f"</tr>"
                )
        if meta_auto_csv_parts:
            auto_csv_items.append(f"[AUTO: TITLE & META ({len(meta_auto_csv_parts)})] " + " || ".join(meta_auto_csv_parts))
        if meta_manual_csv_parts:
            manual_csv_items.append(f"[MANUAL: FLAGGED OVERSIZED TITLE ({len(meta_manual_csv_parts)})] " + " || ".join(meta_manual_csv_parts))

    # C. Visible Body Copy Instances + Awareness & First-Mention CTA Context Bridge Recommendations
    manual_context_cnt = 0
    context_bridge_csv = []
    is_homepage = bool(cat == "Homepage" or re.search(r"^https://www\.servicenow\.com(/(?:uk|au|in|de|fr-ca|fr|es|latam|br|it|nl|jp|kr))?/$", u))

    # 1. Context Bridge for Homepage & Awareness / First-Mention CTA Pages
    if is_homepage:
        manual_context_cnt += 1
        cb_loc = "Hero Marquee Lead Paragraph (Context Bridge)"
        cb_old = "Most companies have AI. Few have the data to ground it, the workflows to deploy it, and the security to govern it. ServiceNow® is the only platform built for all of it, so AI doesn't just think. It acts."
        cb_new = "Most companies have AI but few have the data to ground it, the workflows to deploy it, and the security to govern it. Super Intelligence Control Tower is built for all of it. So AI doesn't just think, it acts."
        cb_note = "Awareness & CTA Context Bridge: Introduces 'Super Intelligence Control Tower' in the marquee body text so context is laid down for visitors before the 'Control Tower' CTA button."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>Awareness Page &amp; First-Mention CTA Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("Awareness & CTA Context Bridge: Introduce 'Super Intelligence Control Tower' in Hero Marquee body copy before 'Control Tower' CTA button")

    elif u.endswith("/products/it-asset-management.html"):
        manual_context_cnt += 1
        cb_loc = "Feature Tab Paragraph Preceding CTA (Context Bridge)"
        cb_old = "Compare AI and machine learning cloud spending across providers. Monitor AI performance metrics to reduce risk and support ethical and regulatory compliance."
        cb_new = "Compare AI and machine learning cloud spending across providers. Monitor AI performance with Super Intelligence Control Tower to reduce risk and support compliance."
        cb_note = "CTA Context Bridge: Introduces 'Super Intelligence Control Tower' in the feature paragraph right before the 'See Control Tower' CTA button."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>First-Mention CTA Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("CTA Context Bridge: Introduce 'Super Intelligence Control Tower' in paragraph preceding 'See Control Tower' button")

    elif u.endswith("/products/context-engine.html"):
        manual_context_cnt += 1
        cb_loc = "Card Paragraph Preceding CTA (Context Bridge)"
        cb_old = "Ground AI decisions in your policies and rules, enforced automatically with explainable audit trails and manual approval when needed."
        cb_new = "Ground AI decisions in your policies and rules with Super Intelligence Control Tower—enforced automatically with explainable audit trails and manual approval when needed."
        cb_note = "CTA Context Bridge: Introduces 'Super Intelligence Control Tower' in the card paragraph right before the 'Explore Control Tower' CTA button."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>First-Mention CTA Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("CTA Context Bridge: Introduce 'Super Intelligence Control Tower' in paragraph preceding 'Explore Control Tower' button")

    elif u.endswith("/platform.html"):
        manual_context_cnt += 1
        cb_loc = "Platform Hub Section Intro Copy (Context Bridge)"
        cb_old = "The AI control tower for business reinvention (Standalone section heading without introductory body paragraph)"
        cb_new = "Connect your data, workflows, and AI on one governed platform with ServiceNow Super Intelligence Control Tower—so AI doesn't just think, it acts."
        cb_note = "Awareness Page Context Bridge: Introduces 'Super Intelligence Control Tower' in the platform lead copy beneath 'The Super Intelligence Control Tower for business reinvention' heading."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>Awareness Page Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("Awareness Page Context Bridge: Introduce 'Super Intelligence Control Tower' in Platform hub section intro copy")

    elif u.endswith("/what-is-servicenow.html"):
        manual_context_cnt += 1
        cb_loc = "Primary Awareness Section Intro Copy (Context Bridge)"
        cb_old = "ServiceNow is the AI control tower for business reinvention—and it all begins with workflows."
        cb_new = "ServiceNow is the Super Intelligence Control Tower for business reinvention—grounding AI in your data, workflows, and governance so AI doesn't just think, it acts."
        cb_note = "Awareness Page Context Bridge: Lay down foundational context for 'Super Intelligence Control Tower' (data, workflows, governance) for top-of-funnel visitors."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>Awareness Page Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("Awareness Page Context Bridge: Expand section intro copy to contextualize 'Super Intelligence Control Tower' for first-time visitors")

    elif u.endswith("/impact.html"):
        manual_context_cnt += 1
        cb_loc = "Impact Feature Lead Copy (Context Bridge)"
        cb_old = "Create a blueprint to drive business outcomes and put the AI control tower for business reinvention to work faster."
        cb_new = "Create a blueprint to drive business outcomes and put ServiceNow Super Intelligence Control Tower—connecting your AI, data, and workflows—to work faster."
        cb_note = "Awareness Page Context Bridge: Contextualizes 'Super Intelligence Control Tower' for visitors on the Impact overview page."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>Awareness Page Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("Awareness Page Context Bridge: Contextualize 'Super Intelligence Control Tower' on Impact overview page")

    elif u.endswith("/events/ai-summits.html"):
        manual_context_cnt += 1
        cb_loc = "Event Hub Section Intro Copy (Context Bridge)"
        cb_old = "Join us to see how ServiceNow acts as the AI control tower and learn how to put AI to work for you."
        cb_new = "Join us to see how ServiceNow Super Intelligence Control Tower grounds AI in your data, workflows, and governance—and learn how to put AI to work for you."
        cb_note = "Awareness Page Context Bridge: Introduces 'Super Intelligence Control Tower' context in the AI Summits hub lead paragraph."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>Awareness Page Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("Awareness Page Context Bridge: Introduce 'Super Intelligence Control Tower' context on AI Summits hub")

    elif u.endswith("/solutions/business-transformation.html"):
        manual_context_cnt += 1
        cb_loc = "Solutions Card Body Copy (Context Bridge)"
        cb_old = "The AI control tower for business reinvention (Standalone card heading without introductory context for Super Intelligence Control Tower)"
        cb_new = "Unite AI, data, and workflows on one governed platform with ServiceNow Super Intelligence Control Tower to accelerate business reinvention."
        cb_note = "Awareness Page Context Bridge: Introduces 'Super Intelligence Control Tower' context beneath the standalone card heading."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>Awareness Page Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("Awareness Page Context Bridge: Add introductory context for 'Super Intelligence Control Tower' beneath card heading")

    elif u.endswith("/products/ai-agents.html") and not body_text_instances:
        manual_context_cnt += 1
        cb_loc = "Section Intro Copy Preceding CTA (Context Bridge)"
        cb_old = "Unify third party AI agents and tools from any platform with ServiceNow Action Fabric. Manage, monitor, and optimize all of them in a single place."
        cb_new = "Unify third party AI agents and tools from any platform with ServiceNow Action Fabric. Manage, monitor, and optimize any AI from a central intelligent hub with ServiceNow Super Intelligence Control Tower."
        cb_note = "CTA Context Bridge: Introduces 'Super Intelligence Control Tower' in the section intro right before the 'Explore Control Tower' CTA button."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>First-Mention CTA Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("CTA Context Bridge: Introduce 'Super Intelligence Control Tower' before 'Explore Control Tower' button")

    elif u.endswith("/products/public-sector-digital-services.html") and not body_text_instances:
        manual_context_cnt += 1
        cb_loc = "Feature Tab Copy Preceding CTA (Context Bridge)"
        cb_old = "Public Sector Digital Services is government's CRM — purpose-built to manage constituent relationships, cases, and service delivery on a single platform. Our solutions connect agency data, workflows, and compliance controls to resolve requests faster and deliver services at scale."
        cb_new = "Public Sector Digital Services is government's CRM — purpose-built to manage constituent relationships, cases, and service delivery on a single platform. Our solutions connect agency data, workflows, and compliance controls to resolve requests faster and deliver services at scale. Running inside the Super Intelligence Control Tower, agencies centralize visibility and control to scale AI confidently across every mission."
        cb_note = "CTA Context Bridge: Introduces 'Super Intelligence Control Tower' in the feature tab copy right before the 'Explore Control Tower' CTA button."
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>First-Mention CTA Context Bridge</div></td>"
            f"<td>{html.escape(cb_old)}</td>"
            f"<td><strong>{html.escape(cb_new)}</strong></td>"
            f"<td>{html.escape(cb_note)}</td>"
            f"</tr>"
        )
        context_bridge_csv.append(f"[MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
        wwsg_csv_notes.append("CTA Context Bridge: Introduce 'Super Intelligence Control Tower' before 'Explore Control Tower' button")

    if body_text_instances:
        auto_body_csv = []
        manual_body_csv = []

        for idx, b in enumerate(body_text_instances, 1):
            b_idx = b.get("instance_num", idx)
            b_loc = b.get("location", "Body Copy")
            b_old = b.get("current_text", "")
            b_new = b.get("target_text", "")

            # In-place Context Bridge paragraphs on ai-agents.html (#2) and public-sector-digital-services.html (#1)
            # where the live paragraph above the 'Explore Control Tower' button does not currently contain 'AI Control Tower'
            if u.endswith("/products/ai-agents.html") and b_idx == 2:
                manual_context_cnt += 1
                cb_loc = "Section Intro Copy Preceding CTA (Context Bridge)"
                cb_old = "Unify third party AI agents and tools from any platform with ServiceNow Action Fabric. Manage, monitor, and optimize all of them in a single place."
                cb_new = "Unify third party AI agents and tools from any platform with ServiceNow Action Fabric. Manage, monitor, and optimize any AI from a central intelligent hub with ServiceNow Super Intelligence Control Tower."
                cb_note = "CTA Context Bridge: Introduces 'Super Intelligence Control Tower' in the section intro right before the 'Explore Control Tower' CTA button."
                drawer_manual_rows.append(
                    f"<tr class='detail-item-manual'>"
                    f"<td><span class='badge b-manual'>Manual</span></td>"
                    f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>First-Mention CTA Context Bridge</div></td>"
                    f"<td>{html.escape(cb_old)}</td>"
                    f"<td><strong>{html.escape(cb_new)}</strong></td>"
                    f"<td>{html.escape(cb_note)}</td>"
                    f"</tr>"
                )
                context_bridge_csv.append(f"#{b_idx} [MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
                wwsg_csv_notes.append(f"Body #{b_idx}: CTA Context Bridge before 'Explore Control Tower' button")
                continue

            if u.endswith("/products/public-sector-digital-services.html") and b_idx == 1:
                manual_context_cnt += 1
                cb_loc = "Feature Tab Copy Preceding CTA (Context Bridge)"
                cb_old = "Public Sector Digital Services is government's CRM — purpose-built to manage constituent relationships, cases, and service delivery on a single platform. Our solutions connect agency data, workflows, and compliance controls to resolve requests faster and deliver services at scale."
                cb_new = "Public Sector Digital Services is government's CRM — purpose-built to manage constituent relationships, cases, and service delivery on a single platform. Our solutions connect agency data, workflows, and compliance controls to resolve requests faster and deliver services at scale. Running inside the Super Intelligence Control Tower, agencies centralize visibility and control to scale AI confidently across every mission."
                cb_note = "CTA Context Bridge: Introduces 'Super Intelligence Control Tower' in the feature tab copy right before the 'Explore Control Tower' CTA button."
                drawer_manual_rows.append(
                    f"<tr class='detail-item-manual'>"
                    f"<td><span class='badge b-manual'>Manual</span></td>"
                    f"<td><strong>{html.escape(cb_loc)}</strong><div class='sub-note'>First-Mention CTA Context Bridge</div></td>"
                    f"<td>{html.escape(cb_old)}</td>"
                    f"<td><strong>{html.escape(cb_new)}</strong></td>"
                    f"<td>{html.escape(cb_note)}</td>"
                    f"</tr>"
                )
                context_bridge_csv.append(f"#{b_idx} [MANUAL_CONTEXT_BRIDGE | {cb_loc}] \"{cb_old}\" -> \"{cb_new}\" ({cb_note})")
                wwsg_csv_notes.append(f"Body #{b_idx}: CTA Context Bridge before 'Explore Control Tower' button")
                continue

            w_res = audit_wwsg_instance(b, u)
            b_loc = w_res.get("display_loc", b_loc)
            if w_res["used_si_short"]:
                si_short_cnt += 1

            if w_res["flags"]:
                fl_str = " ".join(w_res["flags"])
                wwsg_csv_notes.append(f"Body #{b_idx}: {w_res['char_check']} {fl_str}")
            elif w_res["is_over"]:
                wwsg_csv_notes.append(f"Body #{b_idx}: {w_res['char_check']}")

            if w_res["needs_manual_copy"]:
                manual_copy_cnt += 1
                rec_txt = w_res["recommended_auto_target"]
                note_txt = w_res["short_note"]
                drawer_manual_rows.append(
                    f"<tr class='detail-item-manual'>"
                    f"<td><span class='badge b-manual'>Manual</span></td>"
                    f"<td><strong>{html.escape(b_loc)}</strong><div class='sub-note'>Flagged: Name too long — manually decide replacement</div></td>"
                    f"<td>{html.escape(b_old)}</td>"
                    f"<td><strong>{html.escape(rec_txt)}</strong></td>"
                    f"<td>{html.escape(note_txt)}</td>"
                    f"</tr>"
                )
                manual_body_csv.append(f"#{b_idx} [MANUAL_FLAGGED_TOO_LONG | {b_loc}] \"{b_old}\" -> \"{rec_txt}\" ({note_txt})")
            else:
                auto_text_meta_alt_cnt += 1
                rec_auto = w_res["recommended_auto_target"]
                note_txt = w_res["short_note"]
                drawer_auto_rows.append(
                    f"<tr class='detail-item-auto'>"
                    f"<td><span class='badge b-auto'>Auto</span></td>"
                    f"<td><strong>{html.escape(b_loc)}</strong></td>"
                    f"<td>{html.escape(b_old)}</td>"
                    f"<td><strong>{html.escape(rec_auto)}</strong></td>"
                    f"<td>{html.escape(note_txt)}</td>"
                    f"</tr>"
                )
                m_tag = "AUTO_BTN_SHORT" if w_res["used_si_short"] else "AUTO_FULL"
                auto_body_csv.append(f"#{b_idx} [{m_tag} | {b_loc}] \"{b_old}\" -> \"{rec_auto}\"")

        if auto_body_csv:
            auto_csv_items.append(f"[AUTO: BODY INSTANCES ({len(auto_body_csv)})] " + " || ".join(auto_body_csv))
        if manual_body_csv:
            manual_filter_tags.add("m-copy")
            manual_csv_items.append(f"[MANUAL: FLAGGED NAME TOO LONG ({len(manual_body_csv)})] " + " || ".join(manual_body_csv))

    elif body_ai_ct > 0 or raw_loc > 0 or raw_aict > 0:
        cnt = max(1, body_ai_ct + raw_loc + raw_aict)
        auto_text_meta_alt_cnt += cnt
        drawer_auto_rows.append(
            f"<tr class='detail-item-auto'>"
            f"<td><span class='badge b-auto'>Auto</span></td>"
            f"<td><strong>Body Component / Card Copy ({cnt}x)</strong></td>"
            f"<td>AI Control Tower / AICT</td>"
            f"<td><strong>Super Intelligence Control Tower</strong> (or <strong>Control Tower</strong> on CTA buttons)</td>"
            f"<td>Automated text swap</td>"
            f"</tr>"
        )
        auto_csv_items.append(f"[AUTO: BODY COMPONENT COPY ({cnt}x)] Update AI Control Tower / AICT -> 'Super Intelligence Control Tower' (or 'Control Tower' on CTA buttons)")

    if context_bridge_csv:
        manual_filter_tags.add("m-context")
        manual_csv_items.append(f"[MANUAL: CONTEXT BRIDGE ({len(context_bridge_csv)})] " + " || ".join(context_bridge_csv))
    if manual_context_cnt > 0:
        manual_reasons.append(f"Context Bridge ({manual_context_cnt})")
    if manual_copy_cnt > 0:
        manual_reasons.append(f"Name Too Long ({manual_copy_cnt})")

    # D. In-Body Anchor Links & Component Data Attributes (Automated — Deduplicated by href pair)
    if body_slug_links or data_attr_instances or body_slug > 0:
        link_csv_parts = []
        link_groups = []
        link_group_map = {}
        for idx, lk in enumerate(body_slug_links, 1):
            if isinstance(lk, (list, tuple)) and len(lk) >= 3:
                lk_tag, lk_attr, lk_old = lk[0], lk[1], lk[2]
                lk_new = re.sub(r"(?i)ai-control-tower|aict", "super-intelligence-control-tower", lk_old)
            elif isinstance(lk, dict):
                lk_tag = lk.get("tag", "a")
                lk_attr = lk.get("attr", "href")
                lk_old = lk.get("old_href", "")
                lk_new = lk.get("new_href", "")
            else:
                continue
            link_csv_parts.append(f"Link #{idx} [AUTO] (<{lk_tag} {lk_attr}>): {lk_old} -> {lk_new}")
            lkey = (lk_old, lk_new)
            if lkey not in link_group_map:
                lobj = {"old": lk_old, "new": lk_new, "count": 1, "tag": f"<{lk_tag} {lk_attr}>"}
                link_group_map[lkey] = lobj
                link_groups.append(lobj)
            else:
                link_group_map[lkey]["count"] += 1

        for lg in link_groups:
            lk_old = lg["old"]
            lk_new = lg["new"]
            c_str = f" ({lg['count']}x)" if lg["count"] > 1 else ""
            lk_old_clickable = a_link(full_url(lk_old), lk_old) if (lk_old.startswith("/") or lk_old.startswith("http")) else f"<code>{html.escape(lk_old)}</code>"
            drawer_auto_rows.append(
                f"<tr class='detail-item-auto'>"
                f"<td><span class='badge b-auto'>Auto</span></td>"
                f"<td><strong>In-Body Link Href{c_str}</strong><div class='sub-note'><code>{html.escape(lg['tag'])}</code></div></td>"
                f"<td class='mono'>{lk_old_clickable}</td>"
                f"<td class='mono'><code>{html.escape(lk_new)}</code></td>"
                f"<td>Automated link path update</td>"
                f"</tr>"
            )
        if not link_groups and body_slug > 0:
            drawer_auto_rows.append(
                f"<tr class='detail-item-auto'>"
                f"<td><span class='badge b-auto'>Auto</span></td>"
                f"<td><strong>In-Page Slug References ({body_slug}x)</strong></td>"
                f"<td class='mono'><code>.../ai-control-tower...</code></td>"
                f"<td class='mono'><code>.../super-intelligence-control-tower...</code></td>"
                f"<td>Automated link path update</td>"
                f"</tr>"
            )
        if body_slug > 0:
            link_csv_parts.append(f"Total in-page slug refs ({body_slug}x) [AUTO]: ai-control-tower -> super-intelligence-control-tower")
        auto_csv_items.append("[AUTO: LINKS & SLUG REFS] " + " || ".join(link_csv_parts))

    # E. On-Page DAM Image Path Swaps & HTML alt Attributes (100% Automated on Webpages)
    auto_imgs = []
    baked_imgs = []
    for im in imgs:
        if "logo-aicte.svg" in im.lower():
            continue
        fim = full_url(im)
        minfo = MASTER_IMAGES_MAP.get(fim, {})
        if minfo.get("has_baked_aict_pixels", False) or "now-ai-control-tower" in fim or "secure-any-ai-action" in fim:
            if fim not in baked_imgs:
                baked_imgs.append(fim)
        elif "ai-control-tower" in fim.lower() or "aict" in fim.lower():
            if fim not in auto_imgs:
                auto_imgs.append(fim)

    if auto_imgs or baked_imgs or alt_instances:
        img_swap_csv_parts = []
        alt_csv_parts = []
        img_seq = 0
        for im in auto_imgs:
            img_seq += 1
            fim = full_url(im)
            new_fim = get_target_image_url(fim)
            drawer_auto_rows.append(
                f"<tr class='detail-item-auto'>"
                f"<td><span class='badge b-auto'>Auto</span></td>"
                f"<td><strong>DAM Image Path Swap #{img_seq}</strong><div class='sub-note'>0 baked-in pixels (verified by OCR)</div></td>"
                f"<td class='mono'>{a_link(fim)}</td>"
                f"<td class='mono'><code>{html.escape(new_fim)}</code></td>"
                f"<td>Automated image path swap + 301</td>"
                f"</tr>"
            )
            img_swap_csv_parts.append(f"Img #{img_seq} [AUTO_SLUG_SWAP]: {fim} -> {new_fim}")
        for fim in baked_imgs:
            img_seq += 1
            minfo = MASTER_IMAGES_MAP.get(fim, {})
            ocr_disp = clean_ocr_display(minfo.get("matched_ocr_lines", []))
            dims = f"{minfo.get('width', 0)}×{minfo.get('height', 0)}" if minfo.get("width") else "Image"
            kb = f"{minfo.get('bytes', 0)//1024:,} KB" if minfo.get("bytes") else ""
            spec_str = f"{dims}{', ' + kb if kb else ''}"
            new_fim = get_target_image_url(fim)
            target_col_html = (
                f"<div class='mono'><code>{html.escape(new_fim)}</code></div>"
                f"<div class='sub-note' style='margin-top:2px;'>Automated swap to updated graphic copy (<strong>Super Intelligence Control Tower</strong>)</div>"
            )
            drawer_auto_rows.append(
                f"<tr class='detail-item-auto'>"
                f"<td><span class='badge b-auto'>Auto</span></td>"
                f"<td><strong>DAM Image Path Swap #{img_seq}</strong><div class='sub-note'>{html.escape(spec_str)} • Updated graphic copy</div></td>"
                f"<td class='mono'>{a_link(fim)}<div class='sub-note'>OCR: \"{html.escape(ocr_disp)}\"</div></td>"
                f"<td>{target_col_html}</td>"
                f"<td>Automated image path swap (old image &rarr; new image copy)</td>"
                f"</tr>"
            )
            img_swap_csv_parts.append(f"Img #{img_seq} [AUTO_IMAGE_SWAP]: {fim} -> {new_fim} (OCR: \"{ocr_disp}\")")
        for a_idx, a_obj in enumerate(alt_instances, 1):
            a_loc = a_obj.get("location", "<arc-image alt=\"...\">")
            a_old = a_obj.get("current_text", "")
            w_res = audit_wwsg_instance(a_obj, u)
            if w_res["used_si_short"]:
                si_short_cnt += 1
            auto_text_meta_alt_cnt += 1
            best_alt = w_res["recommended_auto_target"]
            drawer_auto_rows.append(
                f"<tr class='detail-item-auto'>"
                f"<td><span class='badge b-auto'>Auto</span></td>"
                f"<td><strong>Image Alt Text #{a_idx}</strong><div class='sub-note'><code>{html.escape(a_loc)}</code></div></td>"
                f"<td>{html.escape(a_old)}</td>"
                f"<td><strong>{html.escape(best_alt)}</strong></td>"
                f"<td>Alt text: {len(best_alt)} chars</td>"
                f"</tr>"
            )
            alt_csv_parts.append(f"Alt #{a_idx} [AUTO] ({len(best_alt)} chars): \"{a_old}\" -> \"{best_alt}\"")
        auto_csv_items.append(
            f"[AUTO: IMAGE PATH SWAPS ({len(auto_imgs) + len(baked_imgs)}) & ALT INSTANCES ({len(alt_instances)})] "
            f"Image Swaps: {' | '.join(img_swap_csv_parts) if img_swap_csv_parts else 'None'} | Alts: {' || '.join(alt_csv_parts) if alt_csv_parts else 'None'}"
        )

    # F. Linked Non-FAQ PDFs with Internal Matches (Manual)
    live_pdfs = [p for p in pdfs if pdf_match_lookup.get(p, (0,))[0] > 0 or "ds-ai-control-tower" in p]
    if live_pdfs:
        manual_reasons.append(f"PDF Update ({len(live_pdfs)})")
        manual_filter_tags.add("m-pdf")
        pdf_csv_parts = []
        for idx, p in enumerate(live_pdfs, 1):
            fp = full_url(p)
            m_info = pdf_match_lookup.get(fp) or pdf_match_lookup.get(p)
            has_pdf_slug = bool("ai-control-tower" in fp.lower() or "aict" in fp.lower())
            if has_pdf_slug:
                new_fp = re.sub(r"(?i)ai-control-tower|aict", "super-intelligence-control-tower", fp)
                pdf_target_html = (
                    f"<div class='mono'><code>{html.escape(new_fp)}</code></div>"
                    f"<div class='sub-note' style='margin-top:2px;'>Update internal PDF text &amp; metadata to <strong>Super Intelligence Control Tower</strong></div>"
                )
                pdf_csv_parts.append(f"{fp} -> {new_fp}")
            else:
                pdf_target_html = "Update internal PDF text &amp; metadata to <strong>Super Intelligence Control Tower</strong>"
                pdf_csv_parts.append(fp)
            if m_info and m_info[0] > 0:
                m_cnt, p_pages, _ = m_info
                p_note = f"{m_cnt} matches across {p_pages} PDF pages • Re-export" + (" + rename filename (HTTP 301)" if has_pdf_slug else "")
            else:
                p_note = "Update datasheet copy + rename filename (HTTP 301)"
            drawer_manual_rows.append(
                f"<tr class='detail-item-manual'>"
                f"<td><span class='badge b-manual'>Manual</span></td>"
                f"<td><strong>Linked PDF Document #{idx}</strong></td>"
                f"<td class='mono'>{a_link(fp)}</td>"
                f"<td>{pdf_target_html}</td>"
                f"<td>{html.escape(p_note)}</td>"
                f"</tr>"
            )
        manual_csv_items.append(f"[MANUAL: LINKED PDFS ({len(live_pdfs)})] " + ", ".join(pdf_csv_parts))

    # G. Localized Flagship Native Grammar Review (Manual L10n Task on 10 non-English flagship pages)
    if loc_code in ("de", "fr", "fr-ca", "es", "latam", "br", "it", "nl", "jp", "kr") and "products/ai-control-tower.html" in u:
        manual_reasons.append("L10n Review")
        manual_filter_tags.add("m-l10n")
        drawer_manual_rows.append(
            f"<tr class='detail-item-manual'>"
            f"<td><span class='badge b-manual'>Manual</span></td>"
            f"<td><strong>Localized Grammar Review ({loc_code.upper()})</strong></td>"
            f"<td>Translated Flagship Product Page</td>"
            f"<td>Verify gender/article agreement and compound noun inflections on translated headings &amp; FAQs</td>"
            f"<td>Post-script L10n QA</td>"
            f"</tr>"
        )
        manual_csv_items.append(f"[MANUAL: L10N GRAMMAR REVIEW ({loc_code.upper()})] Verify gender/article agreement on translated flagship page")

    auto_edit_count = auto_text_meta_alt_cnt + auto_link_cnt + len(auto_imgs) + len(baked_imgs)
    if auto_edit_count == 0:
        auto_edit_count = max(1, r["true_body"] + r.get("raw_localized", 0))

    manual_task_count = (
        (1 if "m-slug" in manual_filter_tags else 0)
        + manual_copy_cnt
        + manual_context_cnt
        + len(live_pdfs)
        + (1 if "m-l10n" in manual_filter_tags else 0)
    )

    all_drawer_rows = "".join(drawer_auto_rows + drawer_manual_rows)
    drawer_table_html = (
        f"<table class='mini-table'>"
        f"<thead><tr>"
        f"<th style='width:8%;'>Track</th>"
        f"<th style='width:22%;'>Element / Location</th>"
        f"<th style='width:27%;'>Current (Live)</th>"
        f"<th style='width:27%;'>New Target (Approved)</th>"
        f"<th style='width:16%;'>Limit &amp; Notes</th>"
        f"</tr></thead>"
        f"<tbody>{all_drawer_rows}</tbody>"
        f"</table>"
    )

    auto_csv_str = " ; ".join(auto_csv_items) if auto_csv_items else "None"
    manual_csv_str = " ; ".join(manual_csv_items) if manual_csv_items else "None (100% Automated)"
    combined_csv_str = " ; ".join(auto_csv_items + manual_csv_items)
    wwsg_summary_str = " | ".join(wwsg_csv_notes) if wwsg_csv_notes else "Product Marketing naming standard: Super Intelligence Control Tower (or Control Tower on CTA buttons)"

    return {
        "is_fully_auto": (manual_task_count == 0),
        "drawer_table_html": drawer_table_html,
        "drawer_row_count": len(drawer_auto_rows) + len(drawer_manual_rows),
        "drawer_auto_row_count": len(drawer_auto_rows),
        "drawer_manual_row_count": len(drawer_manual_rows),
        "auto_csv": auto_csv_str,
        "auto_edit_count": auto_edit_count,
        "auto_text_meta_alt_cnt": auto_text_meta_alt_cnt,
        "auto_link_cnt": auto_link_cnt,
        "auto_img_cnt": len(auto_imgs) + len(baked_imgs),
        "si_short_cnt": si_short_cnt,
        "manual_csv": manual_csv_str,
        "manual_task_count": manual_task_count,
        "manual_copy_cnt": manual_copy_cnt,
        "manual_context_cnt": manual_context_cnt,
        "manual_reasons": manual_reasons,
        "manual_filter_tags": sorted(manual_filter_tags),
        "combined_csv": combined_csv_str,
        "wwsg_summary": wwsg_summary_str,
    }

# ---------------------------------------------------------------------------
# 5. Classify Every Page & Instance into Automated Updates vs. Manual Updates
# ---------------------------------------------------------------------------
wave_summary = {
    "w1": {"label": "Wave 1 (Flagship, Hubs & High-Traffic Pages)", "pages": 0, "us_pages": 0, "reg_pages": 0, "visits": 0, "conv": 0, "instances": 0, "auto_pages": 0, "manual_pages": 0},
    "w2": {"label": "Wave 2 (Commercial Product, Solution & Active GTM)", "pages": 0, "us_pages": 0, "reg_pages": 0, "visits": 0, "conv": 0, "instances": 0, "auto_pages": 0, "manual_pages": 0},
    "w3": {"label": "Wave 3 (Long-Tail Blogs, Archival Editorial & Case Studies)", "pages": 0, "us_pages": 0, "reg_pages": 0, "visits": 0, "conv": 0, "instances": 0, "auto_pages": 0, "manual_pages": 0},
}
cat_rebrand_rollup = {}
locale_counts = {k: 0 for k in LOCALE_BADGE_MAP}
manual_locale_counts = {k: 0 for k in LOCALE_BADGE_MAP}

total_auto_instances = 0
total_manual_copy_instances = 0
total_manual_tasks_all = 0
total_si_short_saves = 0
pages_100pct_auto = 0
pages_needing_manual = 0
us_pages_total = 0
regional_pages_total = 0

for r in unique_body_pages:
    u = full_url(r["url"])
    loc_code = r.get("locale", get_page_locale(u))
    locale_counts[loc_code] = locale_counts.get(loc_code, 0) + 1
    if loc_code == "us":
        us_pages_total += 1
    else:
        regional_pages_total += 1

    an = r["an"]
    wc = r["wave_code"]
    p_inst = PAGE_INSTANCES_MAP.get(u, {})
    inst_cnt = len(p_inst.get("body_text_instances", [])) + len(p_inst.get("meta_instances", [])) + len(p_inst.get("alt_instances", []))
    if inst_cnt == 0:
        inst_cnt = max(1, r["true_body"] + r.get("raw_localized", 0))

    split_info = build_split_page_breakdowns(r)
    r["split_info"] = split_info

    total_auto_instances += split_info["auto_text_meta_alt_cnt"]
    total_manual_copy_instances += split_info["manual_copy_cnt"]
    total_manual_tasks_all += split_info["manual_task_count"]
    total_si_short_saves += split_info["si_short_cnt"]

    if split_info["is_fully_auto"]:
        pages_100pct_auto += 1
        wave_summary[wc]["auto_pages"] += 1
    else:
        pages_needing_manual += 1
        wave_summary[wc]["manual_pages"] += 1
        manual_locale_counts[loc_code] = manual_locale_counts.get(loc_code, 0) + 1

    wave_summary[wc]["pages"] += 1
    if loc_code == "us":
        wave_summary[wc]["us_pages"] += 1
    else:
        wave_summary[wc]["reg_pages"] += 1
    wave_summary[wc]["visits"] += an["ytd_visits"]
    wave_summary[wc]["conv"] += an["ytd_conversions"]
    wave_summary[wc]["instances"] += inst_cnt

    cat = an["category"]
    if cat not in cat_rebrand_rollup:
        cat_rebrand_rollup[cat] = {
            "pages": 0, "us_pages": 0, "reg_pages": 0, "instances": 0, "visits": 0, "conv": 0,
            "vis": an["strategic_visibility"], "urg": an["urgency"],
            "w1": 0, "w2": 0, "w3": 0, "auto_pages": 0, "manual_pages": 0
        }
    cat_rebrand_rollup[cat]["pages"] += 1
    if loc_code == "us":
        cat_rebrand_rollup[cat]["us_pages"] += 1
    else:
        cat_rebrand_rollup[cat]["reg_pages"] += 1
    cat_rebrand_rollup[cat]["instances"] += inst_cnt
    cat_rebrand_rollup[cat]["visits"] += an["ytd_visits"]
    cat_rebrand_rollup[cat]["conv"] += an["ytd_conversions"]
    cat_rebrand_rollup[cat][wc] += 1
    if split_info["is_fully_auto"]:
        cat_rebrand_rollup[cat]["auto_pages"] += 1
    else:
        cat_rebrand_rollup[cat]["manual_pages"] += 1

total_rebrand_visits = sum(ws["visits"] for ws in wave_summary.values())
total_rebrand_conv = sum(ws["conv"] for ws in wave_summary.values())

# ---------------------------------------------------------------------------
# 6. Write Master CSV (With Split Automated vs. Manual Columns)
# ---------------------------------------------------------------------------
csv_rows = []
for r in all_pages:
    u = full_url(r["url"])
    eff = full_url(r.get("effective_url", u))
    src = r.get("source", "sitemap.xml")
    tm = r.get("total_matches", 0)
    loc = r.get("raw_localized", 0)
    tb = r["true_body"]
    ch = r["chrome_est"]
    title = r["title_clean"] or "(Empty / Rate-Limited Tail Fetch)"
    pdfs = ", ".join(full_url(p) for p in r.get("pdfs", []) if "business-continuity-faq" not in p)
    an = r["an"]
    loc_code = r.get("locale_code") or get_page_locale(u)

    if r in unique_body_pages:
        sp = r["split_info"]
        csv_breakdown = sp["combined_csv"]
        auto_csv_bd = sp["auto_csv"]
        manual_csv_bd = sp["manual_csv"]
        auto_cnt_val = sp["auto_edit_count"]
        manual_cnt_val = sp["manual_task_count"]
        wwsg_notes = sp["wwsg_summary"]
        wave = r["wave_lbl"]
        prio = r["prio_lbl"]
        exec_mode = "100% Automated Swap (Groovy Script)" if sp["is_fully_auto"] else ("Hybrid: Automated Bulk Swap + Manual Update (" + ", ".join(sp["manual_reasons"]) + ")")
        sort_k = (r["wave_ord"], 0 if loc_code == "us" else 1, -an["ytd_visits"], -an.get("en_baseline_visits", 0), -an["prio_score"], -(tb + loc))
    elif tm > 0:
        csv_breakdown = f"[AUTO VIA GLOBAL NAV XF ({ch}x)] 0 unique body matches. All {ch} occurrences are in Global Nav XF & Search Prompt JSON — automatically fixed in Wave 1."
        auto_csv_bd = csv_breakdown
        manual_csv_bd = "None (100% Automated via Global Nav XF)"
        auto_cnt_val = ch
        manual_cnt_val = 0
        wwsg_notes = "Global Nav label 'Super Intelligence Control Tower' (32 chars)"
        wave = "Wave 1 (Auto via Global Nav XF)"
        prio = "P2 - Auto-Fixed by Global Nav XF"
        exec_mode = "100% Automated (Global Nav XF)"
        sort_k = (3, 0 if loc_code == "us" else 1, -an["ytd_visits"], 0, 0, u)
    elif r.get("title"):
        csv_breakdown = "[VERIFIED 0 MATCHES] Full HTML fetched; 0 occurrences of AI Control Tower / AICT."
        auto_csv_bd = "None"
        manual_csv_bd = "None"
        auto_cnt_val = 0
        manual_cnt_val = 0
        wwsg_notes = "N/A (0 Matches)"
        wave = "Verified 0 Matches"
        prio = "P3 - Verified 0 Matches"
        exec_mode = "No Action Needed"
        sort_k = (4, 0 if loc_code == "us" else 1, -an["ytd_visits"], 0, 0, u)
    else:
        csv_breakdown = "[RATE-LIMITED TAIL QUEUE] Queued for retry."
        auto_csv_bd = "None"
        manual_csv_bd = "None"
        auto_cnt_val = 0
        manual_cnt_val = 0
        wwsg_notes = "N/A"
        wave = "Retry Queue"
        prio = "P3 - Retry Queue"
        exec_mode = "No Action Needed"
        sort_k = (5, 0 if loc_code == "us" else 1, -an["ytd_visits"], 0, 0, u)

    csv_rows.append({
        "sort_k": sort_k,
        "Wave": wave,
        "Region_Locale": loc_code.upper(),
        "Priority": prio,
        "Execution_Mode": exec_mode,
        "Page_Category": an["category"],
        "YTD_Visits": an["ytd_visits"],
        "EN_Parent_Baseline_Visits": an.get("en_baseline_visits", 0),
        "YTD_Form_Conversions": an["ytd_conversions"],
        "YTD_FCR_Pct": f"{an['ytd_fcr']*100:.2f}%" if an["ytd_fcr"] else "0.00%",
        "YTD_Traffic_Rank": an["ytd_rank"] if an["ytd_rank"] else "",
        "Strategic_Visibility_1_5": an["strategic_visibility"],
        "Urgency_1_3": an["urgency"],
        "Priority_Score": an["prio_score"],
        "Discovery_Source": src,
        "Asset_Type": "Webpage",
        "URL_or_DAM_Path": u,
        "Effective_URL": eff,
        "HTTP_Status": r.get("status", 200) if (tm > 0 or loc > 0 or r.get("title")) else 429,
        "Total_EN_Slug_Matches": tm,
        "Unique_Body_Matches": tb,
        "Global_Nav_And_Search_SSI_Matches": ch,
        "Localized_Term_Matches": loc,
        "Automated_Edits_Count": auto_cnt_val,
        "Manual_Tasks_Count": manual_cnt_val,
        "Page_Title_or_Meta": title,
        "WWSG_Copy_And_Char_Audit": wwsg_notes,
        "Automated_Changes_Breakdown": auto_csv_bd,
        "Manual_Changes_Breakdown": manual_csv_bd,
        "Exact_Page_Change_Breakdown": csv_breakdown,
        "Linked_Non_FAQ_PDFs": pdfs
    })

for p in audited_pdfs:
    if p.get("matches", 0) > 0 or "aict" in p["url"].lower() or "control-tower" in p["url"].lower():
        m_cnt = 6 if "ebk-aict-discovery-playbook" in p["url"] else p.get("matches", 0)
        ocr_note = "[MANUAL: VECTOR PDF RE-EXPORT (6x)] Re-export InDesign source (AICT-DISCOVERY-PLAYBOOK-02) to replace 5 flattened vector text matches on Pages 5, 8, 9 + 1 Meta Title" if "ebk-aict-discovery-playbook" in p["url"] else f"[MANUAL: LIVE DAM PDF ({m_cnt}x)] Replace {m_cnt} occurrences in PDF text & metadata ('{p.get('meta_title', '')}'): " + (" | ".join(p.get("snippets", [])[:2]))
        hosts = pdf_to_hosts.get(full_url(p["url"]), [])
        hosts_full = ", ".join(full_url(h) for h in hosts[:4]) if hosts else full_url(p["url"])
        p_loc = "DE" if "/de-de/" in p["url"] else ("FR" if "/fr-fr/" in p["url"] else "US")
        csv_rows.append({
            "sort_k": (0.5, 0, 0, 0, 0, -m_cnt),
            "Wave": "Wave 1" if "ebk-aict-discovery-playbook" in p["url"] else "Wave 2",
            "Region_Locale": p_loc,
            "Priority": "P0 - Live DAM PDF",
            "Execution_Mode": "Manual Update (PDF Source Re-export)",
            "Page_Category": "Ungated ePDF Pages",
            "YTD_Visits": 0,
            "EN_Parent_Baseline_Visits": 0,
            "YTD_Form_Conversions": 0,
            "YTD_FCR_Pct": "0.00%",
            "YTD_Traffic_Rank": "",
            "Strategic_Visibility_1_5": 3,
            "Urgency_1_3": 1,
            "Priority_Score": 1.90,
            "Discovery_Source": "PDFKit + Vision OCR",
            "Asset_Type": f"DAM PDF ({p.get('pages', 0)} pages, {p.get('bytes', 0)//1024} KB)",
            "URL_or_DAM_Path": full_url(p["url"]),
            "Effective_URL": full_url(p["url"]),
            "HTTP_Status": 200 if p.get("valid_pdf") else 404,
            "Total_EN_Slug_Matches": m_cnt,
            "Unique_Body_Matches": m_cnt,
            "Global_Nav_And_Search_SSI_Matches": 0,
            "Localized_Term_Matches": m_cnt if ("/de-de/" in p["url"] or "/fr-fr/" in p["url"]) else 0,
            "Automated_Edits_Count": 0,
            "Manual_Tasks_Count": m_cnt,
            "Page_Title_or_Meta": p.get("meta_title", ""),
            "WWSG_Copy_And_Char_Audit": "[NAMING: Use 'Super Intelligence Control Tower' in PDF text & metadata]",
        "Automated_Changes_Breakdown": "None",
        "Manual_Changes_Breakdown": ocr_note,
        "Exact_Page_Change_Breakdown": ocr_note,
        "Linked_Non_FAQ_PDFs": hosts_full
    })

for im, host_list in sorted(img_map.items()):
    fim = full_url(im)
    minfo = MASTER_IMAGES_MAP.get(fim, {})
    is_false_pos = "logo-aicte.svg" in fim.lower()
    is_baked = minfo.get("has_baked_aict_pixels", False) or "now-ai-control-tower" in fim or "secure-any-ai-action" in fim
    has_slug = bool(minfo.get("has_aict_slug") or (("ai-control-tower" in fim.lower() or "aict" in fim.lower()) and not is_false_pos))
    if not is_false_pos and not is_baked and not has_slug:
        continue
    hosts_full = ", ".join(full_url(u) for u, _ in host_list[:4]) if host_list else fim
    alts = []
    for _, a_list in host_list:
        for a in a_list:
            if a not in alts:
                alts.append(a)
    alt_str = " | ".join(alts[:2]) if alts else "(0 alt text matches)"
    dims = f"{minfo.get('width', 0)}×{minfo.get('height', 0)}" if minfo.get("width") else "Image"
    kb = f"{minfo.get('bytes', 0)//1024} KB" if minfo.get("bytes") else ""
    asset_type_str = f"DAM Image ({dims}{', ' + kb if kb else ''})" if dims != "Image" else "DAM Image"
    new_fim = get_target_image_url(fim) if not is_false_pos else fim
    if is_false_pos:
        note = "[FALSE POSITIVE — DO NOT RENAME] All India Council for Technical Education (AICTE) partner logo"
        prio = "False Positive - Exclude"
        wave_lbl = "Excluded (False Positive)"
        exec_m = "Exclude (False Positive)"
        auto_c, man_c = 0, 0
        auto_bd, man_bd = "None", "None"
    elif is_baked:
        ocr_disp = clean_ocr_display(minfo.get("matched_ocr_lines", []))
        used_str = f" | Used on {len(host_list)} pages ({hosts_full})" if host_list else ""
        auto_bd = f"[AUTO PAGE IMAGE SWAP] Swap old DAM path {fim} -> {new_fim} across {len(host_list)} host pages + update alt text"
        man_bd = f"[DAM GRAPHIC COPY] Update baked-in pixel text (\"{ocr_disp}\") and upload new copy at {new_fim}"
        note = f"{auto_bd} | {man_bd}{used_str} | Alt: {alt_str}"
        prio = "P0 - Baked-In Image Copy + Auto Swap"
        wave_lbl = "Wave 1"
        exec_m = "Automated Page Swap (After DAM Graphic Copy Upload)"
        auto_c, man_c = 1, 1
    else:
        used_str = f" | Used on {len(host_list)} pages ({hosts_full})" if host_list else ""
        note = f"[AUTO: DAM IMAGE RENAME + PAGE PATH SWAP + ALT TEXT] 0 baked-in pixels; copy/rename {fim} -> {new_fim} (HTTP 301), auto-swap page references & update alt text{used_str}"
        prio = "P1 - DAM Image Slug + Alt Text"
        wave_lbl = "Wave 2"
        exec_m = "100% Automated (DAM Copy/Move + Page Path Swap)"
        auto_c, man_c = 1, 0
        auto_bd, man_bd = note, "None"

    csv_rows.append({
        "sort_k": (0.55 if is_baked else 0.65, 0, 0, 0, 0, -len(host_list)),
        "Wave": wave_lbl,
        "Region_Locale": "GLOBAL",
        "Priority": prio,
        "Execution_Mode": exec_m,
        "Page_Category": "DAM Image Asset",
        "YTD_Visits": 0,
        "EN_Parent_Baseline_Visits": 0,
        "YTD_Form_Conversions": 0,
        "YTD_FCR_Pct": "0.00%",
        "YTD_Traffic_Rank": "",
        "Strategic_Visibility_1_5": 4,
        "Urgency_1_3": 2,
        "Priority_Score": 2.55,
        "Discovery_Source": ", ".join(minfo.get("discovery_sources", ["Hybrid Crawl + Vision OCR"])),
        "Asset_Type": asset_type_str,
        "URL_or_DAM_Path": fim,
        "Effective_URL": fim,
        "HTTP_Status": 200,
        "Total_EN_Slug_Matches": len(host_list),
        "Unique_Body_Matches": len(alts),
        "Global_Nav_And_Search_SSI_Matches": 0,
        "Localized_Term_Matches": 0,
        "Automated_Edits_Count": auto_c,
        "Manual_Tasks_Count": man_c,
        "Page_Title_or_Meta": os.path.basename(fim),
        "WWSG_Copy_And_Char_Audit": "[ACCESSIBILITY: Updated alt text using 'Super Intelligence Control Tower']",
        "Automated_Changes_Breakdown": auto_bd,
        "Manual_Changes_Breakdown": man_bd,
        "Exact_Page_Change_Breakdown": note,
        "Linked_Non_FAQ_PDFs": hosts_full
    })

csv_rows.sort(key=lambda x: x["sort_k"])
fieldnames = [
    "Wave", "Region_Locale", "Priority", "Execution_Mode", "Page_Category", "YTD_Visits",
    "EN_Parent_Baseline_Visits", "YTD_Form_Conversions",
    "YTD_FCR_Pct", "YTD_Traffic_Rank", "Strategic_Visibility_1_5", "Urgency_1_3",
    "Priority_Score", "Discovery_Source", "Asset_Type", "URL_or_DAM_Path",
    "Effective_URL", "HTTP_Status", "Total_EN_Slug_Matches", "Unique_Body_Matches",
    "Global_Nav_And_Search_SSI_Matches", "Localized_Term_Matches",
    "Automated_Edits_Count", "Manual_Tasks_Count", "Page_Title_or_Meta",
    "WWSG_Copy_And_Char_Audit", "Automated_Changes_Breakdown", "Manual_Changes_Breakdown",
    "Exact_Page_Change_Breakdown", "Linked_Non_FAQ_PDFs"
]
with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
    w.writeheader()
    w.writerows(csv_rows)

# ---------------------------------------------------------------------------
# 7. Build Unified Collapsible Page Rows for #pagesTable (Supports All, Auto Only, and Manual Only views)
# ---------------------------------------------------------------------------
pages_table_rows_html = []

manual_wave_counts = {"w1": 0, "w2": 0, "w3": 0}
manual_type_counts = {"m-context": 0, "m-copy": 0, "m-slug": 0, "m-img": 0, "m-pdf": 0, "m-l10n": 0}

for page_idx, r in enumerate(unique_body_pages, 1):
    u = full_url(r["url"])
    title_str = html.escape(r["title_clean"] or "(No Title)")
    an = r["an"]
    wc = r["wave_code"]
    fgrp = map_filter_group(an["category"])
    sp = r["split_info"]
    loc_code = r.get("locale_code") or get_page_locale(u)
    loc_lbl, loc_cls = LOCALE_BADGE_MAP.get(loc_code, (loc_code.upper(), "b-w3"))

    if wc == "w1":
        wave_badge = '<span class="badge b-w1">Wave 1</span>'
    elif wc == "w2":
        wave_badge = '<span class="badge b-w2">Wave 2</span>'
    else:
        wave_badge = '<span class="badge b-w3">Wave 3</span>'

    loc_badge = f'<span class="badge b-loc" style="margin-left:4px;">{loc_lbl}</span>'

    visits_val = an["ytd_visits"]
    conv_val = an["ytd_conversions"]
    fcr_val = an["ytd_fcr"] * 100
    rank_val = an["ytd_rank"]
    p_score = an["prio_score"]
    en_base_v = an.get("en_baseline_visits", 0)
    en_base_p = an.get("en_baseline_prio_score", p_score)

    if visits_val > 0:
        rank_str = f" • Rank #{rank_val:,}" if rank_val else ""
        analytics_html = (
            f"<div class='metric-main'>{visits_val:,} visits</div>"
            f"<div class='sub-note'>{conv_val:,} conv ({fcr_val:.1f}%) • Score {p_score:.2f}{rank_str}</div>"
        )
    elif loc_code != "us":
        parent_str = f"US parent: {en_base_v:,} visits • " if en_base_v > 0 else ""
        analytics_html = (
            f"<div class='metric-reg'>{loc_lbl} Regional Page</div>"
            f"<div class='sub-note'>{parent_str}Score {en_base_p:.2f}</div>"
        )
    else:
        analytics_html = (
            f"<div class='metric-reg'>Included by Match</div>"
            f"<div class='sub-note'>Score {p_score:.2f}</div>"
        )

    auto_copy_cnt = sp["auto_text_meta_alt_cnt"]
    auto_total_cnt = sp["drawer_auto_row_count"]
    man_total_cnt = sp["manual_task_count"]
    link_cnt = sp["auto_link_cnt"]
    img_cnt = sp["auto_img_cnt"]

    auto_sub_parts = []
    if (link_cnt > 0 or img_cnt > 0) and auto_copy_cnt > 0:
        auto_sub_parts.append(f"{auto_copy_cnt} copy")
    if link_cnt > 0:
        auto_sub_parts.append(f"{link_cnt} link{'s' if link_cnt != 1 else ''}")
    if img_cnt > 0:
        auto_sub_parts.append(f"{img_cnt} image swap{'s' if img_cnt != 1 else ''}")
    if sp["si_short_cnt"] > 0:
        auto_sub_parts.append(f"{sp['si_short_cnt']} 'Control Tower' CTA{'s' if sp['si_short_cnt'] != 1 else ''}")
    auto_sub_html = f"<div class='sub-note'>{', '.join(auto_sub_parts)}</div>" if auto_sub_parts else "<div class='sub-note'>Groovy bulk swap</div>"

    auto_col_html = (
        f"<span class='badge b-auto'>{auto_total_cnt} Auto Swap{'s' if auto_total_cnt != 1 else ''}</span>"
        f"{auto_sub_html}"
    )

    if sp["is_fully_auto"]:
        auto_exec_attr = "pure-auto"
        has_manual_attr = "0"
        mtags_attr = ""
        man_col_html = "<span class='badge b-neutral'>0 Manual</span><div class='sub-note'>100% automated</div>"
    else:
        auto_exec_attr = "hybrid-auto"
        has_manual_attr = "1"
        manual_wave_counts[wc] = manual_wave_counts.get(wc, 0) + 1
        for mtag in sp["manual_filter_tags"]:
            manual_type_counts[mtag] = manual_type_counts.get(mtag, 0) + 1
        mtags_attr = " ".join(sp["manual_filter_tags"])
        man_reasons_str = ", ".join(sp["manual_reasons"])
        man_col_html = (
            f"<span class='badge b-manual'>{man_total_cnt} Manual Task{'s' if man_total_cnt != 1 else ''}</span>"
            f"<div class='sub-note'>{html.escape(man_reasons_str)}</div>"
        )

    total_drawer_items = sp["drawer_row_count"]
    auto_drawer_items = sp["drawer_auto_row_count"]
    man_drawer_items = sp["drawer_manual_row_count"]

    pages_table_rows_html.append(f"""
    <tr class="page-row" id="row-{page_idx}" data-row-id="{page_idx}" data-cat="{wc}" data-b30="{fgrp}" data-exec="{auto_exec_attr}" data-has-manual="{has_manual_attr}" data-mtypes="{mtags_attr}" data-loc="{loc_code}" data-total-items="{total_drawer_items}" data-auto-items="{auto_drawer_items}" data-man-items="{man_drawer_items}" onclick="toggleDetail({page_idx}, event)">
      <td data-label="Wave &amp; Region">{wave_badge}{loc_badge}</td>
      <td data-label="Page URL &amp; Title">
        <div class="url-line">{a_link(u)}</div>
        <div class="page-title-line">{title_str} <span class="cat-pill">{html.escape(an["category"])}</span></div>
      </td>
      <td data-label="Traffic (YTD)">{analytics_html}</td>
      <td data-label="Automated" class="col-auto">{auto_col_html}</td>
      <td data-label="Manual" class="col-man">{man_col_html}</td>
      <td data-label="Details" style="text-align:right;white-space:nowrap;">
        <button type="button" class="expand-btn" id="btn-exp-{page_idx}" onclick="event.stopPropagation(); toggleDetail({page_idx}, event)">View Changes ({total_drawer_items})</button>
      </td>
    </tr>
    <tr class="detail-row" id="detail-{page_idx}" style="display:none;">
      <td colspan="6" class="detail-cell">
        <div class="drawer-wrap">
          <div class="drawer-header">
            <div><strong>Page Change Breakdown:</strong> {a_link(u)}</div>
            <div class="sub-note">Showing exact Before &rarr; After changes for the active view mode</div>
          </div>
          {sp["drawer_table_html"]}
        </div>
      </td>
    </tr>""")

# ---------------------------------------------------------------------------
# 8. Build Dedicated 301 Redirects Table Rows (Webpage Slugs + DAM PDFs + DAM Images)
# ---------------------------------------------------------------------------
redirect_rows_html = []
redirect_idx = 0
webpage_redirect_cnt = 0
pdf_redirect_cnt = 0
img_redirect_cnt = 0

seen_redirect_urls = set()
slug_pages_sorted = [r for r in unique_body_pages if "ai-control-tower" in r["url"].lower() or "aict" in r["url"].lower()]

for r in slug_pages_sorted:
    old_u = full_url(r["url"])
    if old_u in seen_redirect_urls:
        continue
    seen_redirect_urls.add(old_u)
    redirect_idx += 1
    webpage_redirect_cnt += 1
    new_u = re.sub(r"ai-control-tower", "super-intelligence-control-tower", old_u, flags=re.I)
    new_u = re.sub(r"aict", "super-intelligence-control-tower", new_u, flags=re.I)
    loc_code = r.get("locale", get_page_locale(old_u))
    loc_lbl, _ = LOCALE_BADGE_MAP.get(loc_code, (loc_code.upper(), "b-w3"))
    wc = r["wave_code"]
    w_badge = '<span class="badge b-w1">Wave 1</span>' if wc == "w1" else ('<span class="badge b-w2">Wave 2</span>' if wc == "w2" else '<span class="badge b-w3">Wave 3</span>')

    redirect_rows_html.append(f"""
    <tr data-rtype="webpage" data-rloc="{loc_code}">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Type &amp; Wave"><span class="badge b-auto">Webpage</span> {w_badge} <span class="badge b-loc">{loc_lbl}</span></td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_u)}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_u)}</code></td>
      <td data-label="Technical Action">
        <span class="badge b-neutral">HTTP 301</span>
        <div class="sub-note" style="margin-top:3px;">Move AEM page node, update canonical &amp; hreflang cluster</div>
      </td>
    </tr>""")

extra_webpage_slugs = [
    ("https://www.servicenow.com/products/ai-control-tower/pricing.html", "us", "w1", "Product Pricing Subpage"),
    ("https://www.servicenow.com/events/on-demand-webinars/put-ai-to-work-with-ai-control-tower.html", "us", "w2", "On-Demand Webinar Landing Page"),
    ("https://www.servicenow.com/lpwbr/meet-ai-control-tower-enterprise-ai-governance-at-machine-speed.html", "us", "w2", "Webinar Promotional Landing Page"),
    ("https://www.servicenow.com/uk/events/live/ai-control-tower.html", "uk", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/au/events/live/ai-control-tower.html", "au", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/in/events/live/ai-control-tower.html", "in", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/de/events/live/ai-control-tower.html", "de", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/fr/events/live/ai-control-tower.html", "fr", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/fr-ca/events/live/ai-control-tower.html", "fr-ca", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/es/events/live/ai-control-tower.html", "es", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/latam/events/live/ai-control-tower.html", "latam", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/br/events/live/ai-control-tower.html", "br", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/it/events/live/ai-control-tower.html", "it", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/nl/events/live/ai-control-tower.html", "nl", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/jp/events/live/ai-control-tower.html", "jp", "w1", "Regional Live Event Landing Page"),
    ("https://www.servicenow.com/kr/events/live/ai-control-tower.html", "kr", "w1", "Regional Live Event Landing Page"),
]
for old_u, loc_code, wc, desc_note in extra_webpage_slugs:
    if old_u in seen_redirect_urls:
        continue
    seen_redirect_urls.add(old_u)
    redirect_idx += 1
    webpage_redirect_cnt += 1
    new_u = re.sub(r"ai-control-tower", "super-intelligence-control-tower", old_u, flags=re.I)
    loc_lbl, _ = LOCALE_BADGE_MAP.get(loc_code, (loc_code.upper(), "b-w3"))
    w_badge = '<span class="badge b-w1">Wave 1</span>' if wc == "w1" else '<span class="badge b-w2">Wave 2</span>'
    redirect_rows_html.append(f"""
    <tr data-rtype="webpage" data-rloc="{loc_code}">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Type &amp; Wave"><span class="badge b-auto">Webpage</span> {w_badge} <span class="badge b-loc">{loc_lbl}</span></td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_u)}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_u)}</code></td>
      <td data-label="Technical Action">
        <span class="badge b-neutral">HTTP 301</span>
        <div class="sub-note" style="margin-top:3px;">{html.escape(desc_note)} — redirect legacy links</div>
      </td>
    </tr>""")

pdf_slug_candidates = []
for p in audited_pdfs:
    fp = full_url(p["url"])
    if ("ai-control-tower" in fp.lower() or "aict" in fp.lower()) and "business-continuity-faq" not in fp.lower():
        pdf_slug_candidates.append(fp)
if legal_pdf_url not in pdf_slug_candidates:
    pdf_slug_candidates.append(legal_pdf_url)
ds_pdf_url = "https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/data-sheet/ds-ai-control-tower.pdf"
if ds_pdf_url not in pdf_slug_candidates:
    pdf_slug_candidates.append(ds_pdf_url)

for old_pdf in pdf_slug_candidates:
    if old_pdf in seen_redirect_urls:
        continue
    seen_redirect_urls.add(old_pdf)
    redirect_idx += 1
    pdf_redirect_cnt += 1
    new_pdf = re.sub(r"ai-control-tower", "super-intelligence-control-tower", old_pdf, flags=re.I)
    new_pdf = re.sub(r"aict", "super-intelligence-control-tower", new_pdf, flags=re.I)
    p_loc = "de" if "/de-de/" in old_pdf else ("fr" if "/fr-fr/" in old_pdf else "us")
    loc_lbl, _ = LOCALE_BADGE_MAP.get(p_loc, ("US", "b-ok"))
    w_badge = '<span class="badge b-w1">Wave 1</span>' if ("ebk-aict-discovery-playbook" in old_pdf or "subscription-unit" in old_pdf or "ds-ai-control-tower" in old_pdf) else '<span class="badge b-w2">Wave 2</span>'
    redirect_rows_html.append(f"""
    <tr data-rtype="pdf" data-rloc="{p_loc}">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Type &amp; Wave"><span class="badge b-manual">DAM PDF</span> {w_badge} <span class="badge b-loc">{loc_lbl}</span></td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_pdf)}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_pdf)}</code></td>
      <td data-label="Technical Action">
        <span class="badge b-neutral">HTTP 301</span>
        <div class="sub-note" style="margin-top:3px;">Publish updated PDF at new DAM path + 301 redirect</div>
      </td>
    </tr>""")

for im in sorted(img_map.keys()):
    if "logo-aicte.svg" in im.lower():
        continue
    old_im = full_url(im)
    minfo = MASTER_IMAGES_MAP.get(old_im, {})
    is_baked = minfo.get("has_baked_aict_pixels", False) or "now-ai-control-tower" in old_im or "secure-any-ai-action" in old_im
    has_slug = "ai-control-tower" in old_im.lower() or "aict" in old_im.lower()
    if not is_baked and not has_slug:
        continue
    if old_im in seen_redirect_urls:
        continue
    seen_redirect_urls.add(old_im)
    redirect_idx += 1
    img_redirect_cnt += 1
    new_im = get_target_image_url(old_im)
    w_badge = '<span class="badge b-w1">Wave 1</span>' if is_baked else '<span class="badge b-w2">Wave 2</span>'
    mode_note = "Upload Updated Graphic Copy + Auto Page Path Swap & 301" if is_baked else "Automated 0-Pixel DAM Copy/Move + Auto Page Swap & 301"
    redirect_rows_html.append(f"""
    <tr data-rtype="image" data-rloc="us">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Type &amp; Wave"><span class="badge b-auto">DAM Image</span> {w_badge}</td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_im)}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_im)}</code></td>
      <td data-label="Technical Action">
        <span class="badge b-neutral">HTTP 301</span>
        <div class="sub-note" style="margin-top:3px;">{mode_note}</div>
      </td>
    </tr>""")

total_redirects_count = redirect_idx

# ---------------------------------------------------------------------------
# 9. Build Category Summary Rows & Media Asset Rows
# ---------------------------------------------------------------------------
cat_summary_rows_html = []
for cat_name, st in sorted(cat_rebrand_rollup.items(), key=lambda x: (-x[1]["visits"], -x[1]["pages"])):
    b_site = cat_baseline_summary.get(cat_name, {})
    site_vis = b_site.get("site_visits", st["visits"])
    site_conv = b_site.get("site_conversions", st["conv"])
    pct_of_cat_visits = (st["visits"] / site_vis * 100) if site_vis > 0 else 0.0
    pct_of_rebrand_visits = (st["visits"] / total_rebrand_visits * 100) if total_rebrand_visits > 0 else 0.0

    wave_dist = []
    if st["w1"] > 0:
        wave_dist.append(f'<span class="badge b-w1">Wave 1: {st["w1"]}</span>')
    if st["w2"] > 0:
        wave_dist.append(f'<span class="badge b-w2">Wave 2: {st["w2"]}</span>')
    if st["w3"] > 0:
        wave_dist.append(f'<span class="badge b-w3">Wave 3: {st["w3"]}</span>')

    cat_summary_rows_html.append(f"""
    <tr>
      <td data-label="Page Category"><strong>{html.escape(cat_name)}</strong></td>
      <td data-label="Rebrand Pages"><strong>{st["pages"]} pages</strong><div class="sub-note">{st["us_pages"]} US + {st["reg_pages"]} Regional</div></td>
      <td data-label="Auto vs. Manual"><span class="badge b-auto">{st["auto_pages"]} Auto Only</span> <span class="badge b-manual">{st["manual_pages"]} Has Manual</span><div class="sub-note">{st["instances"]:,} total instances</div></td>
      <td data-label="US YTD Visits"><strong>{st["visits"]:,}</strong> visits<div class="sub-note">{pct_of_rebrand_visits:.1f}% of rebrand traffic ({pct_of_cat_visits:.1f}% of category)</div></td>
      <td data-label="US Conversions"><strong>{st["conv"]:,}</strong> conv<div class="sub-note">of {site_conv:,} category conversions</div></td>
      <td data-label="Wave Split">{" ".join(wave_dist)}</td>
    </tr>""")

media_rows_html = []

baked_design_groups = {}
zero_pixel_slug_imgs = []
for fim, minfo in sorted(MASTER_IMAGES_MAP.items()):
    if minfo.get("has_baked_aict_pixels", False):
        bkey = re.sub(r"\.(sm|xs|md|lg|xl)(?=\.[a-z0-9]+$)", "", fim, flags=re.I)
        bkey = re.sub(r"-(sm|xs|md|lg|xl)(?=\.[a-z0-9]+$)", "", bkey, flags=re.I)
        baked_design_groups.setdefault(bkey, []).append(minfo)
    elif minfo.get("has_aict_slug", False):
        zero_pixel_slug_imgs.append(minfo)

total_baked_images_count = sum(len(v) for v in baked_design_groups.values())
total_baked_designs_count = len(baked_design_groups)

def _design_sort_key(item):
    bkey, rlist = item
    max_pages = max((x.get("page_count", 0) for x in rlist), default=0)
    is_welcome = 1 if "welcome-ai-control-tower" in bkey else 0
    is_flagship = 1 if ("now-ai-control-tower" in bkey or "secure-any-ai-action" in bkey) else 0
    return (-is_welcome, -is_flagship, -max_pages, bkey)

for d_idx, (bkey, rlist) in enumerate(sorted(baked_design_groups.items(), key=_design_sort_key), 1):
    rlist.sort(key=lambda x: x["url"])
    primary = rlist[0]
    rendition_blocks = []
    for x in rlist:
        old_u = x["url"]
        new_u = get_target_image_url(old_u)
        rendition_blocks.append(
            f"<div style='margin-bottom:4px;'><strong>Current:</strong> {a_link(old_u)}<br>"
            f"<strong>New Copy Target:</strong> <code>{html.escape(new_u)}</code></div>"
        )
    url_links_html = "".join(rendition_blocks)

    all_hosts = []
    for x in rlist:
        for hp in x.get("host_pages", []):
            if hp not in all_hosts:
                all_hosts.append(hp)
    if all_hosts:
        sample_hosts = ", ".join(a_link(hp) for hp in all_hosts[:3])
        more_hosts = f" (+ {len(all_hosts)-3} regional/other pages)" if len(all_hosts) > 3 else ""
        hosts_html = f"<div class='sub-note' style='margin-top:4px;'><strong>Auto-swapped on {len(all_hosts)} page{'s' if len(all_hosts)!=1 else ''}:</strong> {sample_hosts}{more_hosts}</div>"
    else:
        hosts_html = ""

    specs_parts = []
    for x in rlist:
        ext = os.path.splitext(x["url"])[1].upper().lstrip(".") or "IMG"
        specs_parts.append(f"{x.get('width', 0)}×{x.get('height', 0)} {ext} ({x.get('bytes', 0)//1024:,} KB)")
    specs_html = "<br>".join(specs_parts)

    ocr_disp = clean_ocr_display(primary.get("matched_ocr_lines", []))
    slug_badge = '<span class="badge b-auto" style="margin-top:4px;display:inline-block;">Auto Page Path Swap + 301</span>'

    media_rows_html.append(f"""
<tr data-mfilter="baked">
  <td data-label="Track &amp; Group"><span class="badge b-manual">DAM Graphic Copy #{d_idx}</span><div class="sub-note" style="margin-top:4px;">{len(rlist)} file rendition{'s' if len(rlist)!=1 else ''}</div>{slug_badge}</td>
  <td data-label="DAM Asset URL &amp; Host Pages" class="mono">
    {url_links_html}
    {hosts_html}
  </td>
  <td data-label="Specs">{specs_html}</td>
  <td data-label="Detected Text (OCR)"><code>"{html.escape(ocr_disp)}"</code></td>
  <td data-label="Required Action">Update baked-in <code>AI Control Tower</code> pixel text to <strong>Super Intelligence Control Tower</strong> and upload copy at new DAM filename. <strong>Page path swap is 100% Automated</strong> via AEM script + HTTP 301 redirect.</td>
</tr>""")

zero_pixel_pair_lines = []
for x in zero_pixel_slug_imgs:
    old_u = x["url"]
    new_u = get_target_image_url(old_u)
    zero_pixel_pair_lines.append(
        f"<div style='margin-bottom:5px;'><strong>Current:</strong> {a_link(old_u)}<br>"
        f"<strong>New Target:</strong> <code>{html.escape(new_u)}</code></div>"
    )
zero_pixel_links_html = "".join(zero_pixel_pair_lines[:6])
if len(zero_pixel_pair_lines) > 6:
    zero_pixel_links_html += (
        f"<details style='margin-top:6px;'>"
        f"<summary style='cursor:pointer;color:var(--blue);font-weight:600;font-size:11.5px;'>"
        f"View all {len(zero_pixel_pair_lines)} automated DAM filename renames (Current &rarr; New Target)"
        f"</summary>"
        f"<div style='margin-top:6px;padding-top:6px;border-top:1px solid var(--border);'>"
        f"{''.join(zero_pixel_pair_lines[6:])}"
        f"</div></details>"
    )

media_rows_html.append(f"""
<tr data-mfilter="auto-img">
  <td data-label="Track &amp; Group"><span class="badge b-auto">Auto DAM Rename</span><div class="sub-note" style="margin-top:4px;">{len(zero_pixel_slug_imgs)} image files (0 baked pixels)</div><span class="badge b-loc" style="margin-top:4px;display:inline-block;">Rename Filename + 301</span></td>
  <td data-label="DAM Asset URL &amp; Host Pages" class="mono">
    {zero_pixel_links_html}
  </td>
  <td data-label="Specs">PNG / JPG / WebP<br>({len(zero_pixel_slug_imgs)} files)</td>
  <td data-label="Detected Text (OCR)">0 baked-in text matches (verified clean by OCR)</td>
  <td data-label="Required Action"><strong>Zero graphic design work needed.</strong> Script renames DAM filenames (<code>ai-control-tower</code> &rarr; <code>super-intelligence-control-tower</code>) with HTTP 301 redirect and updates HTML <code>alt</code> attributes.</td>
</tr>
<tr data-mfilter="pdf">
  <td data-label="Track &amp; Group"><span class="badge b-manual">Manual PDF</span><div class="sub-note" style="margin-top:4px;">6 OCR / Meta Matches</div><span class="badge b-loc" style="margin-top:4px;display:inline-block;">Rename Filename + 301</span></td>
  <td data-label="DAM Asset URL &amp; Host Pages" class="mono">
    <div><strong>Current:</strong> {a_link('https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/resource-center/e-book/ebk-aict-discovery-playbook.pdf')}<br><strong>New Target:</strong> <code>https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/resource-center/e-book/ebk-super-intelligence-control-tower-discovery-playbook.pdf</code></div>
    <div class="sub-note" style="margin-top:4px;"><strong>Linked from:</strong> {a_link('https://www.servicenow.com/standard/resource-center/ebook/ebk-aict-discovery-playbook.html')}</div>
  </td>
  <td data-label="Specs">11 Pages, PDF<br>(231 MB)</td>
  <td data-label="Detected Text (OCR)">Meta title <code>AICT-DISCOVERY-PLAYBOOK-02</code> + 5 vector text matches on Pages 5, 8, 9</td>
  <td data-label="Required Action">Re-export InDesign source with <strong>Super Intelligence Control Tower</strong>, update metadata title, rename PDF filename, and add HTTP 301 redirect.</td>
</tr>
<tr data-mfilter="pdf">
  <td data-label="Track &amp; Group"><span class="badge b-manual">Manual Legal PDF</span><div class="sub-note" style="margin-top:4px;">7 Text Matches</div><span class="badge b-loc" style="margin-top:4px;display:inline-block;">Rename Filename + 301</span></td>
  <td data-label="DAM Asset URL &amp; Host Pages" class="mono">
    <div><strong>Current:</strong> {a_link('https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/legal/ai-control-tower-servicenow-subscription-unit-overview.pdf')}<br><strong>New Target:</strong> <code>https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/legal/super-intelligence-control-tower-servicenow-subscription-unit-overview.pdf</code></div>
    <div class="sub-note" style="margin-top:4px;"><strong>Linked from:</strong> {a_link('https://www.servicenow.com/products/entitlements-packages.html')}</div>
  </td>
  <td data-label="Specs">4 Pages, PDF<br>(109.7 KB)</td>
  <td data-label="Detected Text (OCR)">Page 1 Title/TOC + Page 3 Section Header &amp; 4 Legal clauses</td>
  <td data-label="Required Action">Coordinate with Legal to publish updated Subscription Unit Overview PDF at renamed DAM filename and configure HTTP 301 redirect.</td>
</tr>
""")

for p in audited_pdfs:
    purl = full_url(p["url"])
    if "ebk-aict-discovery-playbook" in purl:
        continue
    m_cnt = p.get("matches", 0)
    if m_cnt == 0:
        continue
    hosts = pdf_to_hosts.get(purl, [])
    host_links = ", ".join(a_link(h) for h in hosts[:3]) if hosts else a_link("https://www.servicenow.com/standard/resource-center.html")
    sn = html.escape(" | ".join(p.get("snippets", [])[:2]))
    mt = html.escape(p.get("meta_title", "") or "(None)")
    has_pdf_slug = bool("ai-control-tower" in purl.lower() or "aict" in purl.lower())
    if has_pdf_slug:
        new_purl = re.sub(r"(?i)ai-control-tower|aict", "super-intelligence-control-tower", purl)
        pdf_url_block = f"<div><strong>Current:</strong> {a_link(purl)}<br><strong>New Target:</strong> <code>{html.escape(new_purl)}</code></div>"
        pdf_slug_badge = '<span class="badge b-loc" style="margin-top:4px;display:inline-block;">Rename Filename + 301</span>'
        pdf_act_tail = "rename PDF filename (<code>ai-control-tower</code> &rarr; <code>super-intelligence-control-tower</code>), and add HTTP 301 redirect."
    else:
        pdf_url_block = f"<div>{a_link(purl)}</div>"
        pdf_slug_badge = ""
        pdf_act_tail = "and republish at same DAM path."
    media_rows_html.append(f"""
<tr data-mfilter="pdf">
  <td data-label="Track &amp; Group"><span class="badge b-manual">Manual PDF</span><div class="sub-note" style="margin-top:4px;">{m_cnt} Text Matches</div>{pdf_slug_badge}</td>
  <td data-label="DAM Asset URL &amp; Host Pages" class="mono">{pdf_url_block}<div class="sub-note" style="margin-top:4px;"><strong>Linked from:</strong> {host_links}</div></td>
  <td data-label="Specs">{p.get('pages', 0)} Pages, PDF<br>({p.get('bytes', 0)//1024:,} KB)</td>
  <td data-label="Detected Text (OCR)"><strong>Meta Title:</strong> <code>{mt}</code><br><strong>Snippet:</strong> <code>{sn}</code></td>
  <td data-label="Required Action">Update source PDF text &amp; metadata title to <strong>Super Intelligence Control Tower</strong>, {pdf_act_tail}</td>
</tr>""")

# ---------------------------------------------------------------------------
# 10. Write Clean Enterprise Light UI Handoff HTML Report & Markdown Plan
# ---------------------------------------------------------------------------
w1_s = wave_summary["w1"]
w2_s = wave_summary["w2"]
w3_s = wave_summary["w3"]

full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ServiceNow Rebrand Execution Handoff: Super Intelligence Control Tower</title>
<style>
  :root {{
    --bg: #f8fafc;
    --surface: #ffffff;
    --surface-2: #f1f5f9;
    --surface-hover: #f8fafc;
    --border: #e2e8f0;
    --border-strong: #cbd5e1;
    --navy: #032d42;
    --navy-light: #064463;
    --text: #0f172a;
    --muted: #475569;
    --subtle: #64748b;
    --green-text: #15803d;
    --green-bg: #dcfce7;
    --green-border: #86efac;
    --amber-text: #b45309;
    --amber-bg: #fef3c7;
    --amber-border: #fde68a;
    --red-text: #b91c1c;
    --red-bg: #fee2e2;
    --red-border: #fca5a5;
    --blue-text: #0369a1;
    --blue-bg: #e0f2fe;
    --blue-border: #7dd3fc;
    --indigo-text: #4338ca;
    --indigo-bg: #e0e7ff;
    --indigo-border: #a5b4fc;
    --brand-green: #62d84e;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    max-width: 100%;
    overflow-x: hidden;
  }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Inter, Roboto, Helvetica, Arial, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
    padding-bottom: 64px;
    font-size: 13.5px;
    -webkit-font-smoothing: antialiased;
  }}
  a, .ext-link {{
    color: var(--blue-text);
    text-decoration: none;
    font-weight: 500;
    word-break: break-word;
    overflow-wrap: anywhere;
  }}
  a:hover, .ext-link:hover {{
    text-decoration: underline;
    color: var(--navy);
  }}
  header {{
    background: var(--navy);
    color: #ffffff;
    border-bottom: 3px solid var(--brand-green);
    padding: 18px 32px;
  }}
  .header-inner {{
    max-width: 1440px;
    margin: 0 auto;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
  }}
  .eyebrow {{
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--brand-green);
    margin-bottom: 4px;
  }}
  h1 {{
    font-size: 21px;
    font-weight: 700;
    letter-spacing: -0.01em;
    color: #ffffff;
  }}
  .subtitle {{
    color: #cbd5e1;
    font-size: 13px;
    margin-top: 3px;
  }}
  .subtitle a {{
    color: #7dd3fc;
  }}
  .container {{
    max-width: 1440px;
    margin: 20px auto;
    padding: 0 32px;
  }}
  /* KPI Summary Bar */
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 12px;
    margin-bottom: 18px;
  }}
  .kpi-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 16px;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
    cursor: pointer;
    transition: border-color 0.15s, box-shadow 0.15s;
  }}
  .kpi-card:hover {{
    border-color: var(--blue-text);
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.06);
  }}
  .kpi-label {{
    color: var(--muted);
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-weight: 700;
  }}
  .kpi-value {{
    font-size: 22px;
    font-weight: 800;
    margin: 4px 0;
    color: var(--navy);
  }}
  .kpi-value.green {{ color: var(--green-text); }}
  .kpi-value.amber {{ color: var(--amber-text); }}
  .kpi-value.blue {{ color: var(--blue-text); }}
  .kpi-value.indigo {{ color: var(--indigo-text); }}
  .kpi-sub {{
    font-size: 12px;
    color: var(--muted);
    line-height: 1.4;
  }}
  /* Primary Navigation Tabs */
  .tabs {{
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    border-bottom: 2px solid var(--border);
    margin-bottom: 18px;
  }}
  .tab-btn {{
    background: transparent;
    color: var(--muted);
    border: none;
    border-bottom: 3px solid transparent;
    padding: 10px 16px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    margin-bottom: -2px;
    transition: color 0.15s, border-color 0.15s, background 0.15s;
    border-radius: 6px 6px 0 0;
  }}
  .tab-btn:hover {{
    color: var(--navy);
    background: var(--surface-2);
  }}
  .tab-btn.active {{
    color: var(--navy);
    border-bottom-color: var(--navy);
    background: var(--surface);
    font-weight: 700;
  }}
  .tab-count {{
    display: inline-block;
    background: var(--surface-2);
    color: var(--muted);
    border: 1px solid var(--border);
    padding: 1px 6px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
    margin-left: 5px;
  }}
  .tab-btn.active .tab-count {{
    background: var(--navy);
    color: #ffffff;
    border-color: var(--navy);
  }}
  .tab-panel {{ display: none; }}
  .tab-panel.active {{ display: block; }}

  /* Cards & Info Banners */
  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 18px 20px;
    margin-bottom: 18px;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
  }}
  .card-header {{
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    flex-wrap: wrap;
    gap: 8px;
    margin-bottom: 12px;
  }}
  .card h2 {{
    font-size: 16px;
    font-weight: 700;
    color: var(--navy);
  }}
  .card p.desc {{
    color: var(--muted);
    font-size: 13px;
    margin-bottom: 12px;
  }}
  .banner {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-left: 4px solid var(--navy);
    border-radius: 6px;
    padding: 12px 16px;
    margin-bottom: 16px;
    font-size: 13px;
    color: var(--text);
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
  }}
  .banner.green {{ border-left-color: var(--green-text); }}
  .banner.amber {{ border-left-color: var(--amber-text); }}
  .banner.blue {{ border-left-color: var(--blue-text); }}

  /* Compact Filter Bar */
  .filter-bar {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    align-items: center;
    background: var(--surface-2);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px 12px;
    margin-bottom: 14px;
  }}
  .seg-group {{
    display: inline-flex;
    background: var(--surface);
    border: 1px solid var(--border-strong);
    border-radius: 6px;
    overflow: hidden;
  }}
  .seg-btn {{
    background: transparent;
    border: none;
    border-right: 1px solid var(--border);
    padding: 6px 11px;
    font-size: 12px;
    font-weight: 600;
    color: var(--muted);
    cursor: pointer;
  }}
  .seg-btn:last-child {{ border-right: none; }}
  .seg-btn:hover {{ background: var(--surface-2); color: var(--text); }}
  .seg-btn.active {{
    background: var(--navy);
    color: #ffffff;
  }}
  .search-input {{
    flex: 1 1 220px;
    min-width: 180px;
    background: var(--surface);
    border: 1px solid var(--border-strong);
    color: var(--text);
    padding: 6px 11px;
    border-radius: 6px;
    font-size: 12.5px;
  }}
  .search-input:focus, .select-input:focus {{
    outline: 2px solid var(--blue-text);
    outline-offset: -1px;
  }}
  .select-input {{
    background: var(--surface);
    border: 1px solid var(--border-strong);
    color: var(--text);
    padding: 6px 10px;
    border-radius: 6px;
    font-size: 12.5px;
    font-weight: 500;
    cursor: pointer;
  }}
  .action-btn {{
    background: var(--surface);
    border: 1px solid var(--border-strong);
    color: var(--text);
    padding: 6px 10px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
  }}
  .action-btn:hover {{
    background: var(--surface-hover);
    border-color: var(--navy);
  }}
  .results-counter {{
    font-size: 12px;
    font-weight: 600;
    color: var(--muted);
    margin-left: auto;
    white-space: nowrap;
  }}

  /* Tables */
  .table-wrap {{
    width: 100%;
    overflow-x: auto;
    border: 1px solid var(--border);
    border-radius: 8px;
    background: var(--surface);
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }}
  th, td {{
    text-align: left;
    padding: 10px 12px;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
    overflow-wrap: anywhere;
    word-break: break-word;
  }}
  th {{
    background: var(--surface-2);
    color: var(--muted);
    font-weight: 700;
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: 0.04em;
    border-bottom: 1px solid var(--border-strong);
  }}
  tr.page-row {{
    cursor: pointer;
    transition: background 0.1s;
  }}
  tr.page-row:hover td {{
    background: #f8fafc;
  }}
  tr.page-row.expanded td {{
    background: #f1f5f9;
    border-bottom-color: transparent;
  }}
  .page-url-wrap {{
    font-weight: 600;
    font-size: 13px;
  }}
  .sub-meta {{
    font-size: 11.5px;
    color: var(--muted);
    margin-top: 2px;
  }}
  .sub-note {{
    font-size: 11.5px;
    color: var(--muted);
    line-height: 1.4;
  }}
  .expand-btn {{
    background: var(--surface);
    border: 1px solid var(--border-strong);
    color: var(--navy);
    padding: 5px 10px;
    border-radius: 6px;
    font-size: 11.5px;
    font-weight: 600;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 5px;
    white-space: nowrap;
  }}
  .expand-btn:hover {{
    background: var(--navy);
    color: #ffffff;
    border-color: var(--navy);
  }}
  .chevron {{
    display: inline-block;
    font-size: 10px;
    transition: transform 0.15s ease;
  }}
  tr.page-row.expanded .chevron {{
    transform: rotate(90deg);
  }}

  /* Collapsible Detail Drawer & Mini-Table */
  tr.detail-row td.detail-cell {{
    background: #f8fafc;
    padding: 0;
    border-bottom: 2px solid var(--border-strong);
  }}
  .drawer-box {{
    padding: 12px 18px 16px 28px;
  }}
  table.detail-mini-table {{
    width: 100%;
    border: 1px solid var(--border-strong);
    border-radius: 6px;
    background: var(--surface);
    overflow: hidden;
    font-size: 12.5px;
  }}
  table.detail-mini-table th {{
    background: #f1f5f9;
    padding: 7px 10px;
    font-size: 10.5px;
  }}
  table.detail-mini-table td {{
    padding: 8px 10px;
    border-bottom: 1px solid var(--border);
  }}
  table.detail-mini-table tr:last-child td {{
    border-bottom: none;
  }}
  /* View-mode filtering inside opened detail drawers */
  #pagesTable[data-view-mode="auto"] tr.detail-item-manual {{
    display: none !important;
  }}
  #pagesTable[data-view-mode="manual"] tr.detail-item-auto {{
    display: none !important;
  }}

  /* Before / After Diff Lines */
  .diff-line {{
    margin-bottom: 4px;
    line-height: 1.45;
  }}
  .diff-line:last-child {{
    margin-bottom: 0;
  }}
  .diff-label {{
    display: inline-block;
    font-size: 10px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    padding: 1px 5px;
    border-radius: 3px;
    margin-right: 6px;
    vertical-align: middle;
  }}
  .diff-old .diff-label {{
    background: #fee2e2;
    color: #991b1b;
  }}
  .diff-old {{
    color: #475569;
  }}
  .diff-new .diff-label {{
    background: #dcfce7;
    color: #166534;
  }}
  .diff-new {{
    color: #0f172a;
    font-weight: 500;
  }}

  /* Badges */
  .badge {{
    display: inline-block;
    padding: 2px 7px;
    border-radius: 4px;
    font-size: 11px;
    font-weight: 700;
    line-height: 1.35;
    white-space: nowrap;
  }}
  .b-auto, .b-ok {{
    background: var(--green-bg);
    color: var(--green-text);
    border: 1px solid var(--green-border);
  }}
  .b-manual, .b-warn {{
    background: var(--amber-bg);
    color: var(--amber-text);
    border: 1px solid var(--amber-border);
  }}
  .b-w1 {{
    background: var(--red-bg);
    color: var(--red-text);
    border: 1px solid var(--red-border);
  }}
  .b-w2 {{
    background: var(--amber-bg);
    color: var(--amber-text);
    border: 1px solid var(--amber-border);
  }}
  .b-w3, .b-info {{
    background: var(--blue-bg);
    color: var(--blue-text);
    border: 1px solid var(--blue-border);
  }}
  .b-loc {{
    background: var(--indigo-bg);
    color: var(--indigo-text);
    border: 1px solid var(--indigo-border);
  }}
  .b-muted {{
    background: var(--surface-2);
    color: var(--muted);
    border: 1px solid var(--border-strong);
  }}
  .task-chip {{
    display: inline-block;
    background: #fffbeb;
    color: #92400e;
    border: 1px solid #fde68a;
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 11px;
    font-weight: 600;
    margin: 2px 3px 0 0;
  }}

  /* Code & Collapsible Details */
  code, .mono {{
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
    font-size: 11.5px;
    background: #f1f5f9;
    border: 1px solid #e2e8f0;
    padding: 1px 5px;
    border-radius: 4px;
    color: #0f172a;
  }}
  td.mono {{
    background: transparent;
    border: none;
    border-bottom: 1px solid var(--border);
    padding: 10px 12px;
  }}
  details.code-drawer {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    margin-bottom: 16px;
    overflow: hidden;
  }}
  details.code-drawer summary {{
    padding: 12px 16px;
    font-weight: 700;
    font-size: 13.5px;
    color: var(--navy);
    background: var(--surface-2);
    cursor: pointer;
    user-select: none;
  }}
  details.code-drawer summary:hover {{
    background: #e2e8f0;
  }}
  details.code-drawer .drawer-body {{
    padding: 14px 16px;
    border-top: 1px solid var(--border);
  }}
  pre {{
    background: #0f172a;
    color: #f8fafc;
    border-radius: 6px;
    padding: 14px;
    overflow-x: auto;
    white-space: pre-wrap;
    word-break: break-word;
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-size: 11.5px;
    line-height: 1.5;
  }}

  @media (max-width: 1024px) {{
    .container {{ padding: 0 16px; }}
    header {{ padding: 14px 16px; }}
    .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }}
  }}
</style>
<script>
window.currentViewMode = 'all';
window.currentPrimaryTab = 'all';
window.isRestoringUrlState = false;

window.syncUrlState = function(usePushState) {{
  if (window.isRestoringUrlState) return;
  try {{
    var params = new URLSearchParams();
    var tab = window.currentPrimaryTab || 'all';
    if (tab && tab !== 'all') {{
      params.set('tab', tab);
    }}

    if (tab === 'all' || tab === 'auto' || tab === 'manual') {{
      var waveEl = document.getElementById('waveSelect');
      var locEl = document.getElementById('locSelect');
      var scopeEl = document.getElementById('scopeSelect');
      var qEl = document.getElementById('pageSearch');
      if (waveEl && waveEl.value && waveEl.value !== 'all') params.set('wave', waveEl.value);
      if (locEl && locEl.value && locEl.value !== 'all') params.set('loc', locEl.value);
      if (scopeEl && scopeEl.value && scopeEl.value !== 'all') params.set('scope', scopeEl.value);
      if (qEl && qEl.value.trim()) params.set('q', qEl.value.trim());

      var openIds = [];
      document.querySelectorAll('#pagesTable tbody tr.page-row.expanded').forEach(function(r) {{
        if (r.style.display !== 'none') {{
          var idNum = r.getAttribute('data-row-id');
          if (idNum) openIds.push(idNum);
        }}
      }});
      if (openIds.length > 0 && openIds.length <= 25) {{
        params.set('open', openIds.join(','));
      }} else if (openIds.length > 25) {{
        params.set('open', 'visible');
      }}
    }} else if (tab === 'redirects') {{
      var rtEl = document.getElementById('redTypeSelect');
      var rqEl = document.getElementById('redSearch');
      if (rtEl && rtEl.value && rtEl.value !== 'all') params.set('rtype', rtEl.value);
      if (rqEl && rqEl.value.trim()) params.set('rq', rqEl.value.trim());
    }} else if (tab === 'media') {{
      var mtEl = document.getElementById('mediaTypeSelect');
      var mqEl = document.getElementById('mediaSearch');
      if (mtEl && mtEl.value && mtEl.value !== 'all') params.set('mtype', mtEl.value);
      if (mqEl && mqEl.value.trim()) params.set('mq', mqEl.value.trim());
    }}

    var qs = params.toString();
    var newUrl = window.location.pathname + (qs ? '?' + qs : '') + window.location.hash;
    if (usePushState) {{
      window.history.pushState(null, '', newUrl);
    }} else {{
      window.history.replaceState(null, '', newUrl);
    }}
  }} catch (e) {{}}
}};

window.copyShareableLink = function(btnEl) {{
  window.syncUrlState(false);
  var url = window.location.href;
  var done = function() {{
    if (!btnEl) return;
    var orig = btnEl.getAttribute('data-orig-label') || btnEl.textContent;
    btnEl.setAttribute('data-orig-label', orig);
    btnEl.textContent = 'Copied Direct Link!';
    btnEl.style.borderColor = '#15803d';
    btnEl.style.color = '#15803d';
    setTimeout(function() {{
      btnEl.textContent = orig;
      btnEl.style.borderColor = '';
      btnEl.style.color = '';
    }}, 1800);
  }};
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(url).then(done).catch(function() {{
      window.prompt('Copy shareable link:', url);
    }});
  }} else {{
    window.prompt('Copy shareable link:', url);
  }}
}};

window.switchPrimaryTab = function(panelId, viewMode, btnId, skipUrlSync) {{
  document.querySelectorAll('.tab-panel').forEach(function(p) {{ p.classList.remove('active'); }});
  document.querySelectorAll('.tab-btn').forEach(function(b) {{ b.classList.remove('active'); }});
  var panel = document.getElementById(panelId);
  if (panel) panel.classList.add('active');
  var btn = document.getElementById(btnId);
  if (btn) btn.classList.add('active');

  if (panelId === 'tab-pages') {{
    var m = viewMode || 'all';
    window.currentPrimaryTab = m;
    window.setPageViewMode(m, false, true);
  }} else if (panelId === 'tab-redirects') {{
    window.currentPrimaryTab = 'redirects';
  }} else if (panelId === 'tab-media') {{
    window.currentPrimaryTab = 'media';
  }} else if (panelId === 'tab-seo') {{
    window.currentPrimaryTab = 'seo';
  }}

  if (!skipUrlSync) {{
    window.syncUrlState(true);
  }}
}};

window.setPageViewMode = function(mode, updateTopTab, skipUrlSync) {{
  window.currentViewMode = mode;
  window.currentPrimaryTab = mode;
  var tbl = document.getElementById('pagesTable');
  if (tbl) tbl.setAttribute('data-view-mode', mode);

  document.querySelectorAll('.seg-mode-btn').forEach(function(b) {{
    b.classList.toggle('active', b.getAttribute('data-mode') === mode);
  }});

  if (updateTopTab !== false) {{
    document.querySelectorAll('.tab-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    if (mode === 'all') document.getElementById('btn-tab-all').classList.add('active');
    if (mode === 'auto') document.getElementById('btn-tab-auto').classList.add('active');
    if (mode === 'manual') document.getElementById('btn-tab-manual').classList.add('active');
  }}

  var bannerTitle = document.getElementById('pagesBannerTitle');
  var bannerDesc = document.getElementById('pagesBannerDesc');
  if (bannerTitle && bannerDesc) {{
    if (mode === 'all') {{
      bannerTitle.textContent = 'All Updates Combined (Auto + Manual per Page)';
      bannerDesc.textContent = 'Showing all {len(unique_body_pages)} pages. Product Marketing Rule: use "Super Intelligence Control Tower" everywhere (no "SI Control Tower"), use "Control Tower" for buttons, include Manual Context Bridge copy on Awareness & first-mention CTA pages ({manual_type_counts["m-context"]} pages), and flag constrained headings/titles where the full 32-char name is too long.';
    }} else if (mode === 'auto') {{
      bannerTitle.textContent = 'Automated Updates Only ({total_auto_instances:,} Scripted Swaps Across {len(unique_body_pages)} Pages)';
      bannerDesc.textContent = 'Showing only automated DOM, metadata, CTA button ("Control Tower"), link, alt-text, and DAM image path swaps executed by the AEM Groovy script. Manual tasks are hidden inside expanded rows.';
    }} else if (mode === 'manual') {{
      bannerTitle.textContent = 'Manual Updates Queue Only ({pages_needing_manual} Pages with {total_manual_tasks_all} Manual Tasks)';
      bannerDesc.textContent = 'Filtered to the {pages_needing_manual} pages requiring manual action ({manual_type_counts["m-context"]} Context Bridge awareness/CTA pages, {total_manual_copy_instances} headings/titles where "Super Intelligence Control Tower" exceeds 50/60 chars across {manual_type_counts["m-copy"]} pages, 34 page slug moves, 13 PDFs, 10 L10n reviews).';
    }}
  }}

  window.filterPagesTable(true);
  if (!skipUrlSync) {{
    window.syncUrlState(true);
  }}
}};

window.toggleDetail = function(idx, event, skipUrlSync) {{
  if (event && event.target && event.target.closest('a')) return;
  var pRow = document.getElementById('row-' + idx);
  var dRow = document.getElementById('detail-' + idx);
  var btn = document.getElementById('btn-exp-' + idx);
  if (!pRow || !dRow) return;
  var isHidden = (dRow.style.display === 'none' || dRow.style.display === '');
  var totalItems = pRow.getAttribute('data-total-items') || '';
  if (isHidden) {{
    dRow.style.display = 'table-row';
    pRow.classList.add('expanded');
    if (btn) btn.textContent = 'Hide Changes (' + totalItems + ')';
  }} else {{
    dRow.style.display = 'none';
    pRow.classList.remove('expanded');
    if (btn) btn.textContent = 'View Changes (' + totalItems + ')';
  }}
  if (!skipUrlSync) {{
    window.syncUrlState(false);
  }}
}};

window.expandAllVisible = function(expand, skipUrlSync) {{
  var rows = document.querySelectorAll('#pagesTable tbody tr.page-row');
  var expandedCount = 0;
  rows.forEach(function(r) {{
    if (r.style.display === 'none') return;
    var idx = r.id.replace('row-', '');
    var dRow = document.getElementById('detail-' + idx);
    var btn = document.getElementById('btn-exp-' + idx);
    var totalItems = r.getAttribute('data-total-items') || '';
    if (!dRow) return;
    if (expand && expandedCount < 50) {{
      dRow.style.display = 'table-row';
      r.classList.add('expanded');
      if (btn) btn.textContent = 'Hide Changes (' + totalItems + ')';
      expandedCount++;
    }} else if (!expand) {{
      dRow.style.display = 'none';
      r.classList.remove('expanded');
      if (btn) btn.textContent = 'View Changes (' + totalItems + ')';
    }}
  }});
  if (!skipUrlSync) {{
    window.syncUrlState(false);
  }}
}};

window.resetPageFilters = function() {{
  document.getElementById('pageSearch').value = '';
  document.getElementById('waveSelect').value = 'all';
  document.getElementById('locSelect').value = 'all';
  document.getElementById('scopeSelect').value = 'all';
  window.expandAllVisible(false, true);
  window.filterPagesTable(false);
}};

window.filterPagesTable = function(skipUrlSync) {{
  var q = (document.getElementById('pageSearch').value || '').toLowerCase().trim();
  var wave = document.getElementById('waveSelect').value;
  var loc = document.getElementById('locSelect').value;
  var scope = document.getElementById('scopeSelect').value;
  var mode = window.currentViewMode || 'all';

  var rows = document.querySelectorAll('#pagesTable tbody tr.page-row');
  var visible = 0;

  rows.forEach(function(r) {{
    var idx = r.id.replace('row-', '');
    var dRow = document.getElementById('detail-' + idx);
    var btn = document.getElementById('btn-exp-' + idx);
    var totalItems = r.getAttribute('data-total-items') || '';
    var rWave = r.getAttribute('data-cat');
    var rB30 = r.getAttribute('data-b30');
    var rExec = r.getAttribute('data-exec');
    var rLoc = r.getAttribute('data-loc');
    var rMtypes = r.getAttribute('data-mtypes') || '';
    var hasMan = r.getAttribute('data-has-manual') === '1';

    var matchMode = true;
    if (mode === 'manual') {{
      matchMode = hasMan;
    }}

    var matchWave = (wave === 'all') || (rWave === wave);

    var matchLoc = true;
    if (loc === 'us') matchLoc = (rLoc === 'us');
    else if (loc === 'regional') matchLoc = (rLoc !== 'us');
    else if (loc !== 'all') matchLoc = (rLoc === loc);

    var matchScope = true;
    if (scope === 'hybrid-auto' || scope === 'pure-auto') {{
      matchScope = (rExec === scope);
    }} else if (scope.startsWith('m-')) {{
      matchScope = rMtypes.includes(scope);
    }} else if (scope.startsWith('cat-')) {{
      matchScope = (rB30 === scope);
    }}

    var matchQ = true;
    if (q) {{
      var combinedText = (r.innerText + ' ' + (dRow ? dRow.innerText : '')).toLowerCase();
      matchQ = combinedText.includes(q);
    }}

    var show = matchMode && matchWave && matchLoc && matchScope && matchQ;
    r.style.display = show ? '' : 'none';
    if (!show && dRow) {{
      dRow.style.display = 'none';
      r.classList.remove('expanded');
      if (btn) btn.textContent = 'View Changes (' + totalItems + ')';
    }}
    if (show) visible++;
  }});

  var counter = document.getElementById('pagesCounter');
  if (counter) {{
    var totalPool = (mode === 'manual') ? {pages_needing_manual} : {len(unique_body_pages)};
    counter.textContent = 'Showing ' + visible + ' of ' + totalPool + ' pages';
  }}

  if (!skipUrlSync) {{
    window.syncUrlState(false);
  }}
}};

// Tab 4: 301 Redirects Filter
window.filterRedTable = function(skipUrlSync) {{
  var q = (document.getElementById('redSearch').value || '').toLowerCase().trim();
  var rtype = document.getElementById('redTypeSelect').value;
  var rows = document.querySelectorAll('#redTable tbody tr');
  var vis = 0;
  rows.forEach(function(r) {{
    var text = r.innerText.toLowerCase();
    var rt = r.getAttribute('data-rtype');
    var show = (!q || text.includes(q)) && (rtype === 'all' || rt === rtype);
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  var c = document.getElementById('redCounter');
  if (c) c.textContent = 'Showing ' + vis + ' of {total_redirects_count} redirects';
  if (!skipUrlSync) {{
    window.syncUrlState(false);
  }}
}};

// Tab 5: Media & PDF Assets Filter
window.filterMediaTable = function(skipUrlSync) {{
  var q = (document.getElementById('mediaSearch').value || '').toLowerCase().trim();
  var mtype = document.getElementById('mediaTypeSelect').value;
  var rows = document.querySelectorAll('#mediaTable tbody tr');
  var vis = 0;
  rows.forEach(function(r) {{
    var text = r.innerText.toLowerCase();
    var mt = r.getAttribute('data-mfilter');
    var show = (!q || text.includes(q)) && (mtype === 'all' || mt === mtype);
    r.style.display = show ? '' : 'none';
    if (show) vis++;
  }});
  var c = document.getElementById('mediaCounter');
  if (c) c.textContent = 'Showing ' + vis + ' asset rows';
  if (!skipUrlSync) {{
    window.syncUrlState(false);
  }}
}};

window.applyStateFromUrl = function() {{
  window.isRestoringUrlState = true;
  try {{
    var params = new URLSearchParams(window.location.search);
    var tab = params.get('tab') || params.get('mode') || 'all';

    // Reset all controls first
    var ps = document.getElementById('pageSearch');
    var ws = document.getElementById('waveSelect');
    var ls = document.getElementById('locSelect');
    var ss = document.getElementById('scopeSelect');
    if (ps) ps.value = params.get('q') || '';
    if (ws && params.get('wave')) ws.value = params.get('wave'); else if (ws) ws.value = 'all';
    if (ls && params.get('loc')) ls.value = params.get('loc'); else if (ls) ls.value = 'all';
    if (ss && params.get('scope')) ss.value = params.get('scope'); else if (ss) ss.value = 'all';

    var rs = document.getElementById('redSearch');
    var rts = document.getElementById('redTypeSelect');
    if (rs) rs.value = params.get('rq') || '';
    if (rts && params.get('rtype')) rts.value = params.get('rtype'); else if (rts) rts.value = 'all';

    var ms = document.getElementById('mediaSearch');
    var mts = document.getElementById('mediaTypeSelect');
    if (ms) ms.value = params.get('mq') || '';
    if (mts && params.get('mtype')) mts.value = params.get('mtype'); else if (mts) mts.value = 'all';

    if (tab === 'redirects' || tab === 'tab-redirects') {{
      window.switchPrimaryTab('tab-redirects', null, 'btn-tab-redirects', true);
      window.filterRedTable(true);
    }} else if (tab === 'media' || tab === 'tab-media') {{
      window.switchPrimaryTab('tab-media', null, 'btn-tab-media', true);
      window.filterMediaTable(true);
    }} else if (tab === 'seo' || tab === 'tab-seo') {{
      window.switchPrimaryTab('tab-seo', null, 'btn-tab-seo', true);
    }} else {{
      var vMode = (tab === 'auto' || tab === 'manual') ? tab : 'all';
      var bId = (vMode === 'auto') ? 'btn-tab-auto' : ((vMode === 'manual') ? 'btn-tab-manual' : 'btn-tab-all');
      window.switchPrimaryTab('tab-pages', vMode, bId, true);
      window.expandAllVisible(false, true);

      var openParam = params.get('open');
      if (openParam === 'visible' || openParam === 'all') {{
        window.expandAllVisible(true, true);
      }} else if (openParam) {{
        var firstRow = null;
        openParam.split(',').forEach(function(idStr) {{
          var cleanId = idStr.trim();
          if (!cleanId) return;
          var pRow = document.getElementById('row-' + cleanId);
          var dRow = document.getElementById('detail-' + cleanId);
          if (pRow && dRow && pRow.style.display !== 'none') {{
            window.toggleDetail(cleanId, null, true);
            if (!firstRow) firstRow = pRow;
          }}
        }});
        if (firstRow) {{
          setTimeout(function() {{
            firstRow.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
          }}, 100);
        }}
      }}
    }}
  }} catch (e) {{}}
  window.isRestoringUrlState = false;
}};

window.addEventListener('DOMContentLoaded', function() {{
  window.applyStateFromUrl();
}});
window.addEventListener('popstate', function() {{
  window.applyStateFromUrl();
}});
</script>
</head>
<body>

<header>
  <div class="header-inner">
    <div>
      <div class="eyebrow">ServiceNow Digital UX • Global &amp; Regional Rebrand Handoff</div>
      <h1>AI Control Tower &rarr; Super Intelligence Control Tower</h1>
      <div class="subtitle">{len(unique_body_pages)} Total Pages ({us_pages_total} US + {regional_pages_total} Regional Across 13 Locales) • Buttons: "Control Tower" • Zero "SI Control Tower"</div>
    </div>
    <div>
      <button type="button" class="action-btn" onclick="copyShareableLink(this)" title="Copy direct deep-link URL to current tab, filters, and expanded rows" style="font-weight:700;padding:8px 14px;border-color:#cbd5e1;">Copy Direct Link</button>
    </div>
  </div>
</header>

<div class="container">
  <!-- KPI SUMMARY BAR (CLICKABLE SHORTCUTS) -->
  <div class="kpi-grid">
    <div class="kpi-card" onclick="switchPrimaryTab('tab-pages', 'all', 'btn-tab-all')">
      <div class="kpi-label">All Updates (Combined)</div>
      <div class="kpi-value">{len(unique_body_pages)} Pages</div>
      <div class="kpi-sub"><strong>{us_pages_total} US</strong> + <strong>{regional_pages_total} Regional</strong> (13 locales). View Auto + Manual together per page.</div>
    </div>
    <div class="kpi-card" onclick="switchPrimaryTab('tab-pages', 'auto', 'btn-tab-auto')">
      <div class="kpi-label">Automated Swaps</div>
      <div class="kpi-value green">{total_auto_instances:,} Swaps</div>
      <div class="kpi-sub"><strong>{pages_100pct_auto} pages 100% Auto</strong> + {pages_needing_manual} hybrid pages + {len(ssi_only_pages):,} Global Nav XF pages.</div>
    </div>
    <div class="kpi-card" onclick="switchPrimaryTab('tab-pages', 'manual', 'btn-tab-manual')">
      <div class="kpi-label">Manual Updates Queue</div>
      <div class="kpi-value amber">{pages_needing_manual} Pages</div>
      <div class="kpi-sub"><strong>{manual_type_counts['m-context']} Context Bridge</strong>, <strong>{total_manual_copy_instances} oversized headings/titles</strong>, 34 page moves, 13 PDFs, 10 L10n.</div>
    </div>
    <div class="kpi-card" onclick="switchPrimaryTab('tab-redirects', null, 'btn-tab-redirects')">
      <div class="kpi-label">HTTP 301 Redirects</div>
      <div class="kpi-value indigo">{total_redirects_count} Slugs</div>
      <div class="kpi-sub"><strong>{webpage_redirect_cnt} Webpages</strong> + <strong>{pdf_redirect_cnt} DAM PDFs</strong> + <strong>{img_redirect_cnt} DAM Images</strong>.</div>
    </div>
    <div class="kpi-card" onclick="switchPrimaryTab('tab-seo', null, 'btn-tab-seo')">
      <div class="kpi-label">Wave 1 Priority Scope</div>
      <div class="kpi-value blue">96.2% Traffic</div>
      <div class="kpi-sub"><strong>{w1_s['pages']} Wave 1 pages</strong> ({w1_s['us_pages']} US + {w1_s['reg_pages']} Regional • {w1_s['visits']:,} visits).</div>
    </div>
  </div>

  <!-- PRIMARY NAVIGATION TABS -->
  <div class="tabs">
    <button id="btn-tab-all" class="tab-btn active" onclick="switchPrimaryTab('tab-pages', 'all', 'btn-tab-all')">
      All Updates <span class="tab-count">{len(unique_body_pages)} Pages</span>
    </button>
    <button id="btn-tab-auto" class="tab-btn" onclick="switchPrimaryTab('tab-pages', 'auto', 'btn-tab-auto')">
      Automated Only <span class="tab-count">{total_auto_instances:,} Swaps</span>
    </button>
    <button id="btn-tab-manual" class="tab-btn" onclick="switchPrimaryTab('tab-pages', 'manual', 'btn-tab-manual')">
      Manual Only <span class="tab-count">{pages_needing_manual} Pages / {total_manual_tasks_all} Tasks</span>
    </button>
    <button id="btn-tab-redirects" class="tab-btn" onclick="switchPrimaryTab('tab-redirects', null, 'btn-tab-redirects')">
      301 Redirects <span class="tab-count">{total_redirects_count}</span>
    </button>
    <button id="btn-tab-media" class="tab-btn" onclick="switchPrimaryTab('tab-media', null, 'btn-tab-media')">
      Media &amp; PDFs <span class="tab-count">{total_baked_images_count} Baked / 13 PDFs</span>
    </button>
    <button id="btn-tab-seo" class="tab-btn" onclick="switchPrimaryTab('tab-seo', null, 'btn-tab-seo')">
      SEO, Naming &amp; Rollout Rules
    </button>
  </div>

  <!-- ===================================================================== -->
  <!-- PANEL 1: UNIFIED COLLAPSIBLE PAGES TABLE (ALL / AUTOMATED / MANUAL)   -->
  <!-- ===================================================================== -->
  <div id="tab-pages" class="tab-panel active">
    <div class="banner">
      <div>
        <strong id="pagesBannerTitle">All Updates Combined (Auto + Manual per Page)</strong>
        <div id="pagesBannerDesc" style="color:var(--muted);font-size:12.5px;margin-top:2px;">
          Showing all {len(unique_body_pages)} pages. Product Marketing Rule: use "Super Intelligence Control Tower" everywhere (no "SI Control Tower"), use "Control Tower" for buttons, include Manual Context Bridge copy on Awareness &amp; first-mention CTA pages ({manual_type_counts['m-context']} pages), and flag constrained headings/titles where the full 32-char name is too long.
        </div>
      </div>
    </div>

    <!-- SINGLE COMPACT FILTER BAR -->
    <div class="filter-bar">
      <div class="seg-group" role="group" aria-label="View Mode">
        <button type="button" class="seg-btn seg-mode-btn active" data-mode="all" onclick="setPageViewMode('all', true)">All (Auto + Manual)</button>
        <button type="button" class="seg-btn seg-mode-btn" data-mode="auto" onclick="setPageViewMode('auto', true)">Automated Only</button>
        <button type="button" class="seg-btn seg-mode-btn" data-mode="manual" onclick="setPageViewMode('manual', true)">Manual Only ({pages_needing_manual})</button>
      </div>

      <input type="text" id="pageSearch" class="search-input" placeholder="Search URL, title, or Before/After text..." onkeyup="filterPagesTable()">

      <select id="waveSelect" class="select-input" onchange="filterPagesTable()">
        <option value="all">All Waves ({len(unique_body_pages)})</option>
        <option value="w1">Wave 1 ({w1_s['pages']} pages • 96.2% Traffic)</option>
        <option value="w2">Wave 2 ({w2_s['pages']} pages • 3.6% Traffic)</option>
        <option value="w3">Wave 3 ({w3_s['pages']} pages • 0.2% Traffic)</option>
      </select>

      <select id="locSelect" class="select-input" onchange="filterPagesTable()">
        <option value="all">All Regions ({len(unique_body_pages)})</option>
        <option value="us">US Only ({locale_counts['us']})</option>
        <option value="regional">All Regional ({regional_pages_total})</option>
        <option value="uk">UK /uk/ ({locale_counts['uk']})</option>
        <option value="au">Australia /au/ ({locale_counts['au']})</option>
        <option value="in">India /in/ ({locale_counts['in']})</option>
        <option value="de">Germany /de/ ({locale_counts['de']})</option>
        <option value="fr">France /fr/ ({locale_counts['fr']})</option>
        <option value="fr-ca">Canada FR /fr-ca/ ({locale_counts['fr-ca']})</option>
        <option value="es">Spain /es/ ({locale_counts['es']})</option>
        <option value="latam">LATAM /latam/ ({locale_counts['latam']})</option>
        <option value="br">Brazil /br/ ({locale_counts['br']})</option>
        <option value="it">Italy /it/ ({locale_counts['it']})</option>
        <option value="nl">Netherlands /nl/ ({locale_counts['nl']})</option>
        <option value="jp">Japan /jp/ ({locale_counts['jp']})</option>
        <option value="kr">Korea /kr/ ({locale_counts['kr']})</option>
      </select>

      <select id="scopeSelect" class="select-input" onchange="filterPagesTable()">
        <option value="all">All Scope &amp; Task Types</option>
        <optgroup label="Automation vs. Manual Split">
          <option value="hybrid-auto">Has Both Auto + Manual ({pages_needing_manual} pages)</option>
          <option value="pure-auto">100% Automated Only ({pages_100pct_auto} pages)</option>
        </optgroup>
        <optgroup label="Manual Task Type">
          <option value="m-context">Manual: Context Bridge (Awareness &amp; CTA) ({manual_type_counts['m-context']} pages)</option>
          <option value="m-copy">Manual: Name Too Long (&gt;Limit) ({manual_type_counts['m-copy']} pages)</option>
          <option value="m-slug">Manual: 301 URL Slug Move ({manual_type_counts['m-slug']} pages)</option>
          <option value="m-pdf">Manual: Live/Vector PDF Update ({manual_type_counts['m-pdf']} pages)</option>
          <option value="m-l10n">Manual: L10n Grammar Review ({manual_type_counts['m-l10n']} pages)</option>
        </optgroup>
        <optgroup label="Page Category">
          <option value="cat-product">Category: Product Pages</option>
          <option value="cat-solutions">Category: Solutions &amp; Apps</option>
          <option value="cat-platform">Category: Homepage &amp; Platform</option>
          <option value="cat-events">Category: Events &amp; Webinars</option>
          <option value="cat-resources">Category: Resources &amp; ePDFs</option>
          <option value="cat-editorial">Category: Workflow &amp; Blogs</option>
        </optgroup>
      </select>

      <button type="button" class="action-btn" onclick="expandAllVisible(true)" title="Expand first 50 visible rows">Expand Visible</button>
      <button type="button" class="action-btn" onclick="expandAllVisible(false)">Collapse All</button>
      <button type="button" class="action-btn" onclick="resetPageFilters()">Reset</button>
      <button type="button" class="action-btn" onclick="copyShareableLink(this)" title="Copy direct link to this filtered view">Share View</button>

      <span id="pagesCounter" class="results-counter">Showing {len(unique_body_pages)} of {len(unique_body_pages)} pages</span>
    </div>

    <div class="table-wrap">
      <table id="pagesTable" data-view-mode="all">
        <thead>
          <tr>
            <th style="width:11%;">Wave &amp; Region</th>
            <th style="width:34%;">Page URL &amp; Category</th>
            <th style="width:15%;">US Traffic &amp; Basis</th>
            <th style="width:14%;">Automated</th>
            <th style="width:14%;">Manual</th>
            <th style="width:12%;text-align:right;">Details</th>
          </tr>
        </thead>
        <tbody>
          {"".join(pages_table_rows_html)}
        </tbody>
      </table>
    </div>
  </div>

  <!-- ===================================================================== -->
  <!-- PANEL 2: 301 REDIRECTS                                                -->
  <!-- ===================================================================== -->
  <div id="tab-redirects" class="tab-panel">
    <div class="banner">
      <div>
        <strong>HTTP 301 Permanent Redirect Inventory ({total_redirects_count} Total URL &amp; DAM Slug Migrations)</strong>
        <div style="color:var(--muted);font-size:12.5px;margin-top:2px;">
          Covers {webpage_redirect_cnt} Webpage URL slugs (US + 13 regional locales), {pdf_redirect_cnt} DAM PDF filenames, and {img_redirect_cnt} DAM Image filenames.
        </div>
      </div>
    </div>

    <details class="code-drawer">
      <summary>View Ready-to-Deploy Apache Dispatcher (mod_rewrite) &amp; Akamai Edge Redirector Rules</summary>
      <div class="drawer-body">
        <p class="sub-note" style="margin-bottom:8px;">Drop-in configuration handling all US + 13 regional webpage slugs, DAM PDFs (including <code>.coredownload</code>), and DAM images.</p>
        <pre># ServiceNow Apache Dispatcher / Akamai Edge Redirector — 301 Cutover Rules
# Product Rename: AI Control Tower (AICT) -> Super Intelligence Control Tower
RewriteEngine On

# 1. Flagship Product Page, Pricing Subpage + All 13 Regional Locales
RewriteRule ^(/([a-z]{{2}}(-[a-z]{{2}})?|latam))?/products/ai-control-tower(/pricing)?\\.html$ $1/products/super-intelligence-control-tower$4.html [R=301,L,NC]

# 2. Contact Us, Webinar Promo &amp; Discovery Playbook eBook
RewriteRule ^/contact-us-aict\\.html$ /contact-us-super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/lpwbr/meet-ai-control-tower-enterprise-ai-governance-at-machine-speed\\.html$ /lpwbr/meet-super-intelligence-control-tower-enterprise-ai-governance-at-machine-speed.html [R=301,L,NC]
RewriteRule ^/standard/resource-center/ebook/ebk-aict-discovery-playbook\\.html$ /standard/resource-center/ebook/ebk-super-intelligence-control-tower-discovery-playbook.html [R=301,L,NC]

# 3. Live Events, Knowledge Sessions &amp; On-Demand Webinars (US + 13 Regional Locales)
RewriteRule ^(/([a-z]{{2}}(-[a-z]{{2}})?|latam))?/events/live/ai-control-tower\\.html$ $1/events/live/super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/events/on-demand-webinars/put-ai-to-work-with-ai-control-tower\\.html$ /events/on-demand-webinars/put-ai-to-work-with-super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/events/knowledge/2024/sessions/now-on-now--turning-ai-governance-risk-into-reward-with-the-ai-control-tower\\.html$ /events/knowledge/2024/sessions/now-on-now--turning-ai-governance-risk-into-reward-with-the-super-intelligence-control-tower.html [R=301,L,NC]

# 4. Resource Center White Papers, Solution Briefs &amp; eBooks (US + DE/FR/UK)
RewriteRule ^(/([a-z]{{2}}(-[a-z]{{2}})?|latam))?/standard/resource-center/(white-paper|solution-brief|ebook)/(.*)ai-control-tower(.*)\\.html$ $1/standard/resource-center/$4/$5super-intelligence-control-tower$6.html [R=301,L,NC]

# 5. Workflow Articles, Regional Blogs &amp; Documentation Landing Page
RewriteRule ^/workflow/ai/(.*)ai-control-tower\\.html$ /workflow/ai/$1super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/(uk|de|nl)/blogs/2025/(.*)ai-control-tower$ /$1/blogs/2025/$2super-intelligence-control-tower [R=301,L,NC]
RewriteRule ^/docs/r/intelligent-experiences/ai-control-tower/ai-control-tower-landing\\.html$ /docs/r/intelligent-experiences/super-intelligence-control-tower/super-intelligence-control-tower-landing.html [R=301,L,NC]

# 6. DAM PDFs &amp; .coredownload Endpoints (EN-US, DE-DE, FR-FR)
RewriteRule ^/content/dam/servicenow-assets/public/(.*)/ebk-aict-discovery-playbook\\.pdf(\\.coredownload)?$ /content/dam/servicenow-assets/public/$1/ebk-super-intelligence-control-tower-discovery-playbook.pdf$2 [R=301,L,NC]
RewriteRule ^/content/dam/servicenow-assets/public/(.*)/(.*)ai-control-tower(.*)\\.pdf(\\.coredownload)?$ /content/dam/servicenow-assets/public/$1/$2super-intelligence-control-tower$3.pdf$4 [R=301,L,NC]

# 7. DAM Images &amp; Adobe Dynamic Media (.png, .jpg, .webp, .svg — excludes partner logo-aicte.svg)
RewriteCond %{{REQUEST_URI}} !logo-aicte\\.svg$ [NC]
RewriteRule ^/content/dam/servicenow-assets/public/(.*)/(.*)(ai-control-tower|aict)(.*)\\.(png|jpg|jpeg|webp|svg)$ /content/dam/servicenow-assets/public/$1/$2super-intelligence-control-tower$4.$5 [R=301,L,NC]
RewriteRule ^/adobe/assets/urn:aaid:aem:([^/]+)/as/(.*)(ai-control-tower|aict)(.*)\\.(png|jpg|jpeg|webp|svg)$ /adobe/assets/urn:aaid:aem:$1/as/$2super-intelligence-control-tower$4.$5 [R=301,L,NC]</pre>
      </div>
    </details>

    <div class="filter-bar">
      <input type="text" id="redSearch" class="search-input" placeholder="Search old URL, new URL, locale, .pdf, .png..." onkeyup="filterRedTable()">
      <select id="redTypeSelect" class="select-input" onchange="filterRedTable()">
        <option value="all">All Redirect Types ({total_redirects_count})</option>
        <option value="webpage">Webpage URL Slugs ({webpage_redirect_cnt})</option>
        <option value="pdf">DAM PDF Slugs ({pdf_redirect_cnt})</option>
        <option value="image">DAM Image Slugs ({img_redirect_cnt})</option>
      </select>
      <button type="button" class="action-btn" onclick="copyShareableLink(this)" title="Copy direct link to this filtered view">Share View</button>
      <span id="redCounter" class="results-counter">Showing {total_redirects_count} of {total_redirects_count} redirects</span>
    </div>

    <div class="table-wrap">
      <table id="redTable">
        <thead>
          <tr>
            <th style="width:5%;">#</th>
            <th style="width:16%;">Asset Type, Wave &amp; Locale</th>
            <th style="width:33%;">Old Live URL (301 Source)</th>
            <th style="width:30%;">New Target URL (301 Destination)</th>
            <th style="width:16%;">Technical Action</th>
          </tr>
        </thead>
        <tbody>
          {"".join(redirect_rows_html)}
        </tbody>
      </table>
    </div>
  </div>

  <!-- ===================================================================== -->
  <!-- PANEL 3: MEDIA & PDF ASSETS                                           -->
  <!-- ===================================================================== -->
  <div id="tab-media" class="tab-panel">
    <div class="banner">
      <div>
        <strong>DAM Image &amp; PDF Asset Handoff ({total_baked_images_count} Baked-In Image Files Across {total_baked_designs_count} Visual Designs + {total_zero_pixel_images_count} Zero-Pixel Images + 13 PDFs)</strong>
        <div style="color:var(--muted);font-size:12.5px;margin-top:2px;">
          Verified via 4-layer discovery (388 live pages + 8,113 Brand 3.0 DAM assets + responsive variant probing + Apple Vision OCR). All DAM filenames containing <code>ai-control-tower</code> or <code>aict</code> are renamed to <code>super-intelligence-control-tower</code> with HTTP 301 redirects.
        </div>
      </div>
    </div>

    <div class="filter-bar">
      <input type="text" id="mediaSearch" class="search-input" placeholder="Search asset filename, host page URL, or detected OCR text..." onkeyup="filterMediaTable()">
      <select id="mediaTypeSelect" class="select-input" onchange="filterMediaTable()">
        <option value="all">All Media &amp; PDF Assets</option>
        <option value="baked">Manual: Baked-In Pixel Images ({total_baked_designs_count} Design Groups / {total_baked_images_count} Files)</option>
        <option value="pdf">Manual: Live/Vector PDFs (13 Files)</option>
        <option value="auto-img">Automated: Zero-Pixel Filename Moves ({total_zero_pixel_images_count} Files)</option>
      </select>
      <button type="button" class="action-btn" onclick="copyShareableLink(this)" title="Copy direct link to this filtered view">Share View</button>
      <span id="mediaCounter" class="results-counter">Showing all asset groups</span>
    </div>

    <div class="table-wrap">
      <table id="mediaTable">
        <thead>
          <tr>
            <th style="width:14%;">Track &amp; Group</th>
            <th style="width:34%;">DAM Asset URL &amp; Host Pages</th>
            <th style="width:12%;">Specs</th>
            <th style="width:20%;">Detected Text (OCR / PDF)</th>
            <th style="width:20%;">Required Action</th>
          </tr>
        </thead>
        <tbody>
          {"".join(media_rows_html)}
        </tbody>
      </table>
    </div>
  </div>

  <!-- ===================================================================== -->
  <!-- PANEL 4: SEO, NAMING & ROLLOUT RULES                                  -->
  <!-- ===================================================================== -->
  <div id="tab-seo" class="tab-panel">
    <div class="card">
      <div class="card-header">
        <h2>1. Product Marketing Naming Rules: "Super Intelligence Control Tower" (Standard) &amp; "Control Tower" (Buttons)</h2>
      </div>
      <p class="desc">Per Product Marketing direction, do <strong>not</strong> use <code>SI Control Tower</code> anywhere. All copy, headings, metadata, and former <code>AICT</code> acronyms transition to <strong><code>Super Intelligence Control Tower</code></strong> (32 chars). CTA buttons use <strong><code>Control Tower</code></strong> (13 chars). Body copy character-limit trims are skipped to minimize manual work; constrained headings (&gt;50 chars) and compound <code>&lt;title&gt;</code> tags (&gt;60 chars) where <code>Super Intelligence Control Tower</code> is too long are flagged in the Manual queue so the team can manually decide what to replace them with.</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th style="width:22%;">Placement / Component</th>
              <th style="width:12%;">Char Limit</th>
              <th style="width:22%;">Approved Target Form</th>
              <th style="width:12%;">Track</th>
              <th style="width:32%;">Exact Target Copy &amp; Rationale</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Flagship Product <code>&lt;title&gt;</code> &amp; Social Titles</strong><br>{a_link('https://www.servicenow.com/products/ai-control-tower.html')}</td>
              <td>&le; 60 chars</td>
              <td><span class="badge b-auto">Super Intelligence Control Tower</span></td>
              <td><span class="badge b-auto">Auto</span></td>
              <td><code>Super Intelligence Control Tower - ServiceNow</code> (45/60 chars). Establishes canonical entity for Google SERPs and AI Overviews.</td>
            </tr>
            <tr>
              <td><strong>Compound <code>&lt;title&gt;</code> Tags</strong><br>(Blogs, Webinars, Solution Briefs)</td>
              <td>&le; 60 chars</td>
              <td><span class="badge b-manual">Flagged if &gt; 60 chars</span></td>
              <td><span class="badge b-manual">Manual</span> (if &gt;60)</td>
              <td>Replaced with <code>Super Intelligence Control Tower</code>; any compound <code>&lt;title&gt;</code> that exceeds 60 chars (e.g., <code>What Is a Super Intelligence Control Tower? - ServiceNow</code> [56/60 fits] vs longer titles) is flagged in the Manual queue for editorial decision.</td>
            </tr>
            <tr>
              <td><strong>Flagship Hero Marquee <code>&lt;h1&gt;</code></strong></td>
              <td>&le; 50 chars</td>
              <td><span class="badge b-auto">Super Intelligence Control Tower</span></td>
              <td><span class="badge b-auto">Auto</span></td>
              <td><code>Super Intelligence Control Tower</code> (32/50 chars). Direct 1:1 swap with zero legacy product eyebrow text.</td>
            </tr>
            <tr>
              <td><strong>Body Copy, Marquee Lead Paragraphs &amp; FAQ Answers</strong></td>
              <td>Skipped for this update</td>
              <td><span class="badge b-auto">Super Intelligence Control Tower</span></td>
              <td><span class="badge b-auto">Auto</span></td>
              <td>100% automated direct replacement of <code>AI Control Tower</code> and <code>AICT</code> with <code>Super Intelligence Control Tower</code>. WWSG body copy trimming is skipped for this update to reduce manual work.</td>
            </tr>
            <tr>
              <td><strong>CTA Buttons &amp; Awareness Context Bridge Standard</strong><br>{a_link('https://www.servicenow.com/')}<br>{a_link('https://www.servicenow.com/platform.html')}<br>{a_link('https://www.servicenow.com/what-is-servicenow.html')}</td>
              <td>CTA: &le; 25 chars<br>Body: &le; 175–250</td>
              <td><span class="badge b-auto">Control Tower</span> (Buttons)<br><span class="badge b-manual">Context Bridge</span> (Lead Copy)</td>
              <td><span class="badge b-auto">Auto CTA</span> +<br><span class="badge b-manual">Manual Context</span></td>
              <td>
                Per Product Marketing direction, CTA buttons use <code>Control Tower</code>:<br>
                &bull; <code>See AI Control Tower</code> &rarr; <strong><code>See Control Tower</code></strong> (17/25 chars)<br>
                &bull; <code>Explore AI Control Tower</code> &rarr; <strong><code>Explore Control Tower</code></strong> (21/25 chars)<br>
                &bull; <code>Extend AI Control Tower</code> &rarr; <strong><code>Extend Control Tower</code></strong> (20/25 chars)<br>
                <strong>Awareness &amp; First-Mention CTA Context Bridge Standard ({manual_type_counts['m-context']} pages):</strong> Whenever a page is an awareness-categorized entry point (<code>Homepage</code>, <code>What is ServiceNow</code>, <code>Platform - Main Page</code>, <code>Impact Pages</code>, <code>Events - AI Summits Hub Page</code>, tagline-only <code>Solutions</code> pages) or uses a short <code>Control Tower</code> CTA button without prior context, a <strong>Manual Context Bridge</strong> copy update is queued to introduce <strong><code>Super Intelligence Control Tower</code></strong> in the surrounding Marquee/Feature paragraph (e.g., Homepage Marquee: <em>&ldquo;Most companies have AI but few have the data to ground it, the workflows to deploy it, and the security to govern it. Super Intelligence Control Tower is built for all of it. So AI doesn't just think, it acts.&rdquo;</em>).
              </td>
            </tr>
            <tr>
              <td><strong>Company Positioning Tagline</strong><br>(Homepages &amp; Platform Hubs)</td>
              <td>&le; 175 chars</td>
              <td><span class="badge b-auto">Super Intelligence Control Tower</span></td>
              <td><span class="badge b-auto">Auto</span></td>
              <td><code>ServiceNow is the Super Intelligence Control Tower for business reinvention</code> (75/175 chars) and <code>The Super Intelligence Control Tower for business reinvention</code> (61/175 chars). Always use full name and capitalize <strong>Control Tower</strong> (Title Case required even where live copy had lowercase <code>AI control tower</code>).</td>
            </tr>
            <tr>
              <td><strong>Constrained Headings, Cards &amp; FAQ Questions (&gt;50 Chars)</strong></td>
              <td>Heading: &le; 50</td>
              <td><span class="badge b-manual">Flagged when &gt; 50 chars</span></td>
              <td><span class="badge b-manual">Manual</span></td>
              <td><strong>{total_manual_copy_instances} total instances across {manual_type_counts['m-copy']} pages</strong> are flagged as <code>Manual: Name Too Long</code> where <code>Super Intelligence Control Tower</code> causes a heading, FAQ question title, tab label, or <code>&lt;title&gt;</code> to exceed its limit so the team can manually decide the replacement.</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div class="card-header">
        <h2>2. Wave Rollout Summary ({len(unique_body_pages)} Pages: {us_pages_total} US + {regional_pages_total} Regional)</h2>
      </div>
      <p class="desc">Regional pages automatically inherit the Wave assignment of their US parent page so global and regional launches stay synchronized.</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th style="width:10%;">Wave</th>
              <th style="width:44%;">Included Page Groups</th>
              <th style="width:22%;">Pages &amp; Execution Split</th>
              <th style="width:12%;">US YTD Visits</th>
              <th style="width:12%;">US Conversions</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><span class="badge b-w1">Wave 1</span></td>
              <td>Global Nav Experience Fragment ({len(ssi_only_pages):,} pages) + Flagship Product &amp; Playbook Pages (US + 13 locales) + US &amp; 13 Regional Homepages ({a_link('https://www.servicenow.com/')}, {a_link('https://www.servicenow.com/uk/')}, etc.) + High-Traffic Hubs (&ge;10,000 visits or &ge;250 conversions).</td>
              <td><strong>{w1_s['pages']} Pages</strong> ({w1_s['us_pages']} US + {w1_s['reg_pages']} Reg)<br><span class="sub-note">{w1_s['auto_pages']} Pure Auto / {w1_s['manual_pages']} Hybrid Manual</span></td>
              <td><strong>{w1_s['visits']:,}</strong><br><span class="badge b-auto">96.2% Share</span></td>
              <td><strong>{w1_s['conv']:,}</strong><br><span class="badge b-auto">90.8% Share</span></td>
            </tr>
            <tr>
              <td><span class="badge b-w2">Wave 2</span></td>
              <td>All remaining Commercial Product, Solution, Capability, Application, Platform, Industry, Event, Resource Center, and active Workflow pages across US + 13 locales.</td>
              <td><strong>{w2_s['pages']} Pages</strong> ({w2_s['us_pages']} US + {w2_s['reg_pages']} Reg)<br><span class="sub-note">{w2_s['auto_pages']} Pure Auto / {w2_s['manual_pages']} Hybrid Manual</span></td>
              <td><strong>{w2_s['visits']:,}</strong><br><span class="badge b-manual">3.6% Share</span></td>
              <td><strong>{w2_s['conv']:,}</strong><br><span class="badge b-manual">8.9% Share</span></td>
            </tr>
            <tr>
              <td><span class="badge b-w3">Wave 3</span></td>
              <td>Long-tail archival Blogs, low-traffic Workflow articles (&lt;100 visits), Customer Stories, Partner Finder, University, and Docs across US + 13 locales.</td>
              <td><strong>{w3_s['pages']} Pages</strong> ({w3_s['us_pages']} US + {w3_s['reg_pages']} Reg)<br><span class="sub-note">{w3_s['auto_pages']} Pure Auto / {w3_s['manual_pages']} Hybrid Manual</span></td>
              <td><strong>{w3_s['visits']:,}</strong><br><span class="badge b-w3">0.2% Share</span></td>
              <td><strong>{w3_s['conv']:,}</strong><br><span class="badge b-w3">0.3% Share</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <div class="card-header">
        <h2>3. Regional Localization Matrix (13 Locales • {regional_pages_total} Regional Pages)</h2>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th style="width:18%;">Region / Locale</th>
              <th style="width:12%;">Pages</th>
              <th style="width:28%;">Key URLs (Homepage &amp; Flagship)</th>
              <th style="width:18%;">Current Live Term</th>
              <th style="width:24%;">Approved Target Full Name &amp; Button Form</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>UK (<code>/uk/</code>), AU (<code>/au/</code>), IN (<code>/in/</code>)</strong></td>
              <td>{locale_counts['uk']} UK / {locale_counts['au']} AU / {locale_counts['in']} IN</td>
              <td class="mono">{a_link('https://www.servicenow.com/uk/')}<br>{a_link('https://www.servicenow.com/uk/products/ai-control-tower.html')}</td>
              <td><code>AI Control Tower</code></td>
              <td><strong>Super Intelligence Control Tower</strong><br><span class="sub-note">Button: <code>Control Tower</code></span></td>
            </tr>
            <tr>
              <td><strong>Germany (<code>/de/</code>)</strong></td>
              <td>{locale_counts['de']} Pages</td>
              <td class="mono">{a_link('https://www.servicenow.com/de/')}<br>{a_link('https://www.servicenow.com/de/products/ai-control-tower.html')}</td>
              <td><code>KI-Kontrollturm</code></td>
              <td><strong>Super-Intelligenz-Kontrollturm</strong><br><span class="sub-note">Button: <code>Kontrollturm</code></span></td>
            </tr>
            <tr>
              <td><strong>France (<code>/fr/</code>) &amp; Canada FR (<code>/fr-ca/</code>)</strong></td>
              <td>{locale_counts['fr']} FR / {locale_counts['fr-ca']} FR-CA</td>
              <td class="mono">{a_link('https://www.servicenow.com/fr/')}<br>{a_link('https://www.servicenow.com/fr/products/ai-control-tower.html')}</td>
              <td><code>Tour de contrôle IA</code></td>
              <td><strong>Tour de contrôle Super Intelligence</strong><br><span class="sub-note">Button: <code>Tour de contrôle</code></span></td>
            </tr>
            <tr>
              <td><strong>Spain (<code>/es/</code>) &amp; LATAM (<code>/latam/</code>)</strong></td>
              <td>{locale_counts['es']} ES / {locale_counts['latam']} LATAM</td>
              <td class="mono">{a_link('https://www.servicenow.com/es/')}<br>{a_link('https://www.servicenow.com/es/products/ai-control-tower.html')}</td>
              <td><code>Torre de control de IA</code></td>
              <td><strong>Torre de control de Superinteligencia</strong><br><span class="sub-note">Button: <code>Torre de control</code></span></td>
            </tr>
            <tr>
              <td><strong>Brazil (<code>/br/</code>)</strong></td>
              <td>{locale_counts['br']} Pages</td>
              <td class="mono">{a_link('https://www.servicenow.com/br/')}<br>{a_link('https://www.servicenow.com/br/products/ai-control-tower.html')}</td>
              <td><code>Torre de controle de IA</code></td>
              <td><strong>Torre de controle de Superinteligência</strong><br><span class="sub-note">Button: <code>Torre de controle</code></span></td>
            </tr>
            <tr>
              <td><strong>Italy (<code>/it/</code>)</strong></td>
              <td>{locale_counts['it']} Pages</td>
              <td class="mono">{a_link('https://www.servicenow.com/it/')}<br>{a_link('https://www.servicenow.com/it/products/ai-control-tower.html')}</td>
              <td><code>Torre di controllo AI</code></td>
              <td><strong>Torre di controllo della Super Intelligenza</strong><br><span class="sub-note">Button: <code>Torre di controllo</code></span></td>
            </tr>
            <tr>
              <td><strong>Netherlands (<code>/nl/</code>)</strong></td>
              <td>{locale_counts['nl']} Pages</td>
              <td class="mono">{a_link('https://www.servicenow.com/nl/')}<br>{a_link('https://www.servicenow.com/nl/products/ai-control-tower.html')}</td>
              <td><code>AI-controletoren</code></td>
              <td><strong>Superintelligentie-controletoren</strong><br><span class="sub-note">Button: <code>Controletoren</code></span></td>
            </tr>
            <tr>
              <td><strong>Japan (<code>/jp/</code>)</strong></td>
              <td>{locale_counts['jp']} Pages</td>
              <td class="mono">{a_link('https://www.servicenow.com/jp/')}<br>{a_link('https://www.servicenow.com/jp/products/ai-control-tower.html')}</td>
              <td><code>AI コントロールタワー</code></td>
              <td><strong>スーパーインテリジェンス・コントロールタワー</strong><br><span class="sub-note">Button: <code>コントロールタワー</code></span></td>
            </tr>
            <tr>
              <td><strong>Korea (<code>/kr/</code>)</strong></td>
              <td>{locale_counts['kr']} Pages</td>
              <td class="mono">{a_link('https://www.servicenow.com/kr/')}<br>{a_link('https://www.servicenow.com/kr/products/ai-control-tower.html')}</td>
              <td><code>AI 컨트롤 타워</code></td>
              <td><strong>슈퍼 인텔리전스 컨트롤 타워</strong><br><span class="sub-note">Button: <code>컨트롤 타워</code></span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <details class="code-drawer">
      <summary>View 31-Category Page Breakdown Table ({len(unique_body_pages)} Pages)</summary>
      <div class="drawer-body">
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Page Category</th>
                <th>Pages (US + Regional)</th>
                <th>Automated vs. Manual Split</th>
                <th>US YTD Visits (% Share)</th>
                <th>US Conversions</th>
                <th>Wave Split</th>
              </tr>
            </thead>
            <tbody>
              {"".join(cat_summary_rows_html)}
            </tbody>
          </table>
        </div>
      </div>
    </details>

    <details class="code-drawer">
      <summary>View Ready-to-Run AEM Groovy Console Script ({total_auto_instances:,} Automated Swaps Across US + 13 Locales)</summary>
      <div class="drawer-body">
        <pre>// AEM Groovy Console Script — Global + Regional Automated Swap ("Super Intelligence Control Tower" + "Control Tower" for CTA Buttons)
import javax.jcr.Node

boolean dryRun = true // Set to false after reviewing dry-run output
String[] paths = [
    "/content/experience-fragments/servicenow",
    "/content/servicenow",
    "/content/dam/servicenow-assets/public"
]

int updatedProps = 0
paths.each {{ rootPath -&gt;
    def rootNode = session.getNode(rootPath)
    rootNode.recurse {{ Node node -&gt;
        node.properties.each {{ prop -&gt;
            if (!prop.isMultiple() &amp;&amp; prop.type == javax.jcr.PropertyType.STRING) {{
                String val = prop.string
                String orig = val
                String pName = prop.name.toLowerCase()

                if (!val.contains("logo-aicte.svg")) {{
                    val = val.replaceAll("(?i)/products/ai-control-tower", "/products/super-intelligence-control-tower")
                    val = val.replaceAll("(?i)contact-us-aict", "contact-us-super-intelligence-control-tower")
                    val = val.replaceAll("(?i)ebk-aict-discovery-playbook", "ebk-super-intelligence-control-tower-discovery-playbook")
                    // Automated DAM Image Path Swap: swap old image path to new DAM image copy path
                    if (val.contains("/content/dam/") || val.contains("/adobe/assets/")) {{
                        val = val.replaceAll("(?i)ai-control-tower|aict", "super-intelligence-control-tower")
                        val = val.replaceAll("(?i)secure-any-ai-action(\\\\.(?:sm|xs|md|lg|xl))?\\\\.(jpg|png|webp)", "secure-any-ai-action-super-intelligence-control-tower\$1.\$2")
                    }}
                }}

                boolean isButton = pName.contains("cta") || pName.contains("button") || pName.contains("linktext") || pName.contains("action") ||
                    (orig.trim().length() &lt;= 42 &amp;&amp; orig.trim().matches("(?i)^(See|Explore|Extend|Discover|Try|Get|Watch|View|Learn about|Infos zu|Voir|D[\\u00e9e]couvrir|Ver|Explorar|Veja|Scopri|Bekijk)\\\\s+.*|.*(\\u3092\\u898b\\u308b|\\ubcf4\\uae30)$"))

                if (isButton) {{
                    val = val
                        .replaceAll("(?i)\\\\bServiceNow AI Control Tower\\\\b", "Control Tower")
                        .replaceAll("(?i)\\\\bAI Control Tower\\\\b", "Control Tower")
                        .replaceAll("(?i)\\\\bAICT\\\\b", "Control Tower")
                        .replaceAll("KI-Kontrollturm", "Kontrollturm")
                        .replaceAll("(?i)Tour de contr\\\\u00f4le (?:de l'IA|IA)", "Tour de contr\\\\u00f4le")
                        .replaceAll("(?i)Torre de control de IA", "Torre de control")
                        .replaceAll("(?i)Torre de controle de IA", "Torre de controle")
                        .replaceAll("(?i)Torre di controllo (?:dell'IA|AI|IA)", "Torre di controllo")
                        .replaceAll("(?i)AI-controletoren", "Controletoren")
                        .replaceAll("AI\\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc|AI \\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc", "\\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc")
                        .replaceAll("AI \\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc|AI \\ucee8\\ud2b8\\ub864\\ud0c0\\uc6cc", "\\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc")
                }} else {{
                    val = val
                        .replaceAll("(?i)\\\\bAI Control Tower\\\\b", "Super Intelligence Control Tower")
                        .replaceAll("(?i)\\\\bAICT\\\\b", "Super Intelligence Control Tower")
                        .replaceAll("KI-Kontrollturm", "Super-Intelligenz-Kontrollturm")
                        .replaceAll("(?i)Tour de contr\\\\u00f4le (?:de l'IA|IA)", "Tour de contr\\\\u00f4le Super Intelligence")
                        .replaceAll("(?i)Torre de control de IA", "Torre de control de Superinteligencia")
                        .replaceAll("(?i)Torre de controle de IA", "Torre de controle de Superintelig\\\\u00eancia")
                        .replaceAll("(?i)Torre di controllo (?:dell'IA|AI|IA)", "Torre di controllo della Super Intelligenza")
                        .replaceAll("(?i)AI-controletoren", "Superintelligentie-controletoren")
                        .replaceAll("AI\\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc|AI \\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc", "\\u30b9\\u30fc\\u30d1\\u30fc\\u30a4\\u30f3\\u30c6\\u30ea\\u30b8\\u30a7\\u30f3\\u30b9\\u30fb\\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc")
                        .replaceAll("AI \\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc|AI \\ucee8\\ud2b8\\ub864\\ud0c0\\uc6cc", "\\uc288\\ud37c \\uc778\\ud154\\ub9ac\\uc804\\uc2a4 \\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc")
                }}

                // Flag constrained headings (>50 chars) or SEO titles (>60 chars) for manual decision
                int maxCap = 0
                if (pName.contains("jcr:title") || pName.contains("navtitle") || pName.contains("heading")) maxCap = 50
                else if (pName.contains("pagetitle") || pName.contains("seotitle")) maxCap = 60

                if (val != orig) {{
                    if (maxCap &gt; 0 &amp;&amp; val.length() &gt; maxCap &amp;&amp; !orig.toLowerCase().contains("business reinvention")) {{
                        println "[FLAG-MANUAL-LENGTH &gt;${{maxCap}}] ${{node.path}}@${{prop.name}} (${{val.length()}} chars): ${{val}}"
                    }} else {{
                        updatedProps++
                        println "[UPDATE] ${{node.path}}@${{prop.name}} (${{val.length()}} chars)"
                        if (!dryRun) prop.setValue(val)
                    }}
                }}
            }}
        }}
    }}
}}
if (!dryRun) session.save()
println "Total properties updated: ${{updatedProps}} (dryRun=${{dryRun}})"</pre>
      </div>
    </details>

    <details class="code-drawer">
      <summary>View JSON-LD Structured Data Schema (SoftwareApplication)</summary>
      <div class="drawer-body">
        <pre>&lt;script type="application/ld+json"&gt;
{{
  "@context": "https://schema.org",
  "@graph": [
    {{
      "@type": "SoftwareApplication",
      "@id": "https://www.servicenow.com/products/super-intelligence-control-tower.html#software",
      "name": "ServiceNow Super Intelligence Control Tower",
      "alternateName": [
        "Super Intelligence Control Tower"
      ],
      "applicationCategory": "BusinessApplication",
      "ApplicationSubCategory": "Enterprise AI Governance, Observability, and Management Platform",
      "operatingSystem": "ServiceNow AI Platform",
      "url": "https://www.servicenow.com/products/super-intelligence-control-tower.html",
      "description": "Enforce runtime governance for AI at scale. Super Intelligence Control Tower gives you a central hub to discover, secure, govern, observe, and measure enterprise AI.",
      "brand": {{ "@type": "Brand", "name": "ServiceNow" }}
    }}
  ]
}}
&lt;/script&gt;</pre>
      </div>
    </details>
  </div>

</div>

</body>
</html>
"""

open(OUT_HTML, "w", encoding="utf-8").write(full_html)

OUT_MD = "/Users/tamer.mina/.gemini/antigravity-cli/brain/37667ee2-9f2b-4c38-a3b2-645e77963657/aict_rebrand_plan.md"
plan_md = f"""# Super Intelligence Control Tower — Global & Regional Rebrand Handoff Plan

> [!IMPORTANT]
> **Product Marketing Naming & Automation Direction Applied Across All {len(unique_body_pages)} Pages (US + 13 Regional Locales + 63 Audited DAM Images)**
> - **Standard Product Name**: **`Super Intelligence Control Tower`** (32 chars). We do **not** use `SI Control Tower` anywhere—all copy, headings, metadata, and former `AICT` acronyms transition to `Super Intelligence Control Tower`.
> - **CTA Button Name**: **`Control Tower`** (13 chars) on buttons (e.g., `See Control Tower` [17/25 chars], `Explore Control Tower` [21/25 chars], `Extend Control Tower` [20/25 chars], and localized button equivalents).
> - **Awareness & First-Mention CTA Context Bridge Standard (`Manual: Context Bridge`)**: Across **{manual_type_counts['m-context']} pages** (all awareness-categorized pages—`Homepage`, `What is ServiceNow`, `Platform - Main Page`, `Impact Pages`, `Events - AI Summits Hub Page`, and tagline-only `Solutions` pages—plus first-mention `Control Tower` CTA button pages across US + 13 locales), a **Manual Context Bridge** copy update is queued to introduce `Super Intelligence Control Tower` in the surrounding Marquee/Feature paragraph (e.g., Homepage Marquee: *"Most companies have AI but few have the data to ground it, the workflows to deploy it, and the security to govern it. Super Intelligence Control Tower is built for all of it. So AI doesn't just think, it acts."*).
> - **All On-Page Image Swaps Are Automated (`Auto`)**: Once updated image copies are uploaded to the new DAM paths (`super-intelligence-control-tower`), the AEM Groovy automation script swaps out the old image paths and puts in the new image paths across all webpages automatically.
> - **WWSG Body Copy Edits Skipped**: Body copy, Marquee lead descriptions, and FAQ answers are 100% automated direct swaps (`AI Control Tower` -> `Super Intelligence Control Tower`) to minimize manual work.
> - **Oversized Headings & Titles Flagged for Manual Decision**: **{total_manual_copy_instances} instances across {manual_type_counts['m-copy']} pages** are flagged as `Manual: Name Too Long` where `Super Intelligence Control Tower` causes a constrained heading (`<h2>`/`<h3>`/`<h4>`, card heading, FAQ question title, or tab label >50 chars) or compound `<title>` (>60 chars) to exceed its character limit.
> - **Interactive Dashboard**: [aict-rebrand-interactive-report.html](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-interactive-report.html)
> - **Master CSV Inventory**: [aict-rebrand-inventory.csv](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-inventory.csv)

---

## 1. Automated vs. Manual Execution Split

| Track & Owner | Scope | Action Summary |
| :--- | :--- | :--- |
| **Track A1: Automated Global Nav XF**<br>*Owner: AEM Platform Admin* | **2 Centralized XFs** ({len(ssi_only_pages):,} pages) | Single Experience Fragment update swaps `"AI Control Tower"` to `"Super Intelligence Control Tower"` and updates the nav link across US + 13 locales. |
| **Track A2: Automated AEM Groovy Swap**<br>*Owner: AEM Engineering* | **All {len(unique_body_pages)} Pages** ({pages_100pct_auto} Pure Auto + {pages_needing_manual} Hybrid)<br>**{total_auto_instances:,} Automated Swaps** | Updates `<title>`, `<meta>`, `<h1>` Hero headings (`Super Intelligence Control Tower`, 32/50 chars), all body copy & FAQ answers, CTA buttons (`See Control Tower`, `Explore Control Tower`, `Extend Control Tower`), company taglines (`ServiceNow is the Super Intelligence Control Tower for business reinvention`), internal links, **all on-page DAM image paths (`old image -> new image copy`)**, and image `alt` tags. |
| **Track M0: Awareness & CTA Context Bridge Copy**<br>*Owner: Digital UX Copywriter / PMM* | **{manual_type_counts['m-context']} Pages** (US + 13 locales) | Update Marquee/Feature lead paragraphs on awareness-categorized pages (`Homepage`, `Platform`, `What is ServiceNow`, `Impact`, `AI Summits`, `Solutions`) and first-mention `Control Tower` CTA button pages so `Super Intelligence Control Tower` is explicitly introduced in context. |
| **Track M1: Flagged Oversized Headings & `<title>` Tags**<br>*Owner: Digital UX Copywriter / PMM* | **{total_manual_copy_instances} Flagged Instances** across **{manual_type_counts['m-copy']} pages** | Review constrained headings (>50 chars) and compound `<title>` tags (>60 chars) where `"Super Intelligence Control Tower"` is too long and decide the replacement text. |
| **Track M2: Manual AEM Move & 301 Redirects**<br>*Owner: AEM Authoring & Web Ops* | **{total_redirects_count} Total 301 Redirects** ({webpage_redirect_cnt} Webpages + {pdf_redirect_cnt} PDFs + {img_redirect_cnt} Images) | Move AEM page nodes (`/products/ai-control-tower.html` -> `/products/super-intelligence-control-tower.html`) and deploy Apache Dispatcher / Akamai 301 rules. |
| **Track M3: DAM Graphic Copy Upload & PDF Updates**<br>*Owner: Creative, Docs & Legal* | **{total_baked_images_count} Baked-In Image Copies ({total_baked_designs_count} Visual Designs)** + **13 PDFs** | Upload updated graphic copies at new `super-intelligence-control-tower` DAM filenames (page swaps are 100% automated) and update 13 live/vector PDFs. |
| **Track M4: Manual Localization QA**<br>*Owner: Regional Web Leads* | **10 Non-English Flagship Pages** | Linguistic review of gender/article agreement on DE, FR, FR-CA, ES, LATAM, BR, IT, NL, JP, and KR flagship product pages. |
"""

open(OUT_MD, "w", encoding="utf-8").write(plan_md)
print(f"Successfully generated clean enterprise light handoff dashboard {OUT_HTML}, {OUT_CSV}, and {OUT_MD}")
print(f"Total Body-Match Pages: {len(unique_body_pages)} (US: {us_pages_total} + Regional: {regional_pages_total})")
print(f"Pages 100% Auto: {pages_100pct_auto} | Pages Needing Manual: {pages_needing_manual} | Auto Instances: {total_auto_instances} | Context Bridge Pages: {manual_type_counts['m-context']} | Flagged Name-Too-Long Instances: {total_manual_copy_instances} (across {manual_type_counts['m-copy']} pages) | Manual Tasks All: {total_manual_tasks_all} | Button CTAs ('Control Tower'): {total_si_short_saves}")
print(f"Total 301 Redirects: {total_redirects_count} (Webpages: {webpage_redirect_cnt}, PDFs: {pdf_redirect_cnt}, Images: {img_redirect_cnt})")




