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
    "us": ("🇺🇸 US", "b-ok"),
    "uk": ("🇬🇧 UK", "b-w3"),
    "au": ("🇦🇺 AU", "b-w3"),
    "in": ("🇮🇳 IN", "b-w3"),
    "de": ("🇩🇪 DE", "b-warn"),
    "fr": ("🇫🇷 FR", "b-warn"),
    "fr-ca": ("🇨🇦 FR-CA", "b-warn"),
    "es": ("🇪🇸 ES", "b-warn"),
    "latam": ("🌎 LATAM", "b-warn"),
    "br": ("🇧🇷 BR", "b-warn"),
    "it": ("🇮🇹 IT", "b-warn"),
    "nl": ("🇳🇱 NL", "b-warn"),
    "jp": ("🇯🇵 JP", "b-warn"),
    "kr": ("🇰🇷 KR", "b-warn"),
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

def a_link(path_or_url, label=None, color="var(--green)"):
    href = full_url(path_or_url)
    disp = label if label is not None else href
    return (
        f'<a class="ext-link" href="{html.escape(href)}" target="_blank" rel="noopener noreferrer" '
        f'title="Open {html.escape(href)} in browser" '
        f'style="color:{color};text-decoration:underline;text-underline-offset:2px;word-break:break-word;overflow-wrap:anywhere;cursor:pointer;">'
        f'{html.escape(disp)} ↗</a>'
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
# 2. WWSG 2026 Q3 Copy Engine + "SI Control Tower" vs "Super Intelligence Control Tower"
#    + Automated vs. Manual Update Classifier (Clean Cutover — Zero Legacy Eyebrows)
# ---------------------------------------------------------------------------
BUZZWORDS_MAP = [
    (r"\bleveraging\b", "using", "leveraging → using"),
    (r"\bleverage\b", "use", "leverage → use"),
    (r"\butilizing\b", "using", "utilizing → using"),
    (r"\butilize\b", "use", "utilize → use"),
    (r"\boptimizing\b", "improving", "optimizing → improving"),
    (r"\boptimize\b", "improve", "optimize → improve"),
    (r"\baccelerating\b", "speeding up", "accelerating → speeding up"),
    (r"\baccelerate\b", "speed up", "accelerate → speed up"),
    (r"\brobust\b", "powerful", "robust → powerful"),
    (r"\bseamlessly\b", "smoothly", "seamlessly → smoothly"),
    (r"\bseamless\b", "smooth", "seamless → smooth"),
    (r"\bsingle pane of glass\b", "central hub", "single pane of glass → central hub"),
    (r"\bempowering\b", "helping", "empowering → helping"),
    (r"\bempowers\b", "helps", "empowers → helps"),
    (r"\bempower\b", "help", "empower → help"),
]

ABSOLUTE_MAP = [
    (r"\bevery AI agent, model, and identity\b", "AI agents, models, and identities", "every AI agent... → AI agents..."),
    (r"\bevery AI asset\b", "AI assets", "every AI asset → AI assets"),
    (r"\bevery AI use case\b", "AI use cases", "every AI use case → AI use cases"),
    (r"\bevery AI initiative\b", "AI initiatives", "every AI initiative → AI initiatives"),
    (r"\bevery cloud and vendor\b", "clouds and vendors", "every cloud and vendor → clouds and vendors"),
    (r"\bevery use case\b", "all use cases", "every use case → enterprise use cases"),
    (r"\balways\b", "consistently", "always → consistently"),
    (r"\bensures\b", "helps ensure", "ensures → helps ensure"),
    (r"\bensure\b", "help ensure", "ensure → help ensure"),
    (r"\bcompletely\b", "thoroughly", "completely → thoroughly"),
    (r"\bentirely\b", "fully", "entirely → fully"),
    (r"\beliminate\b", "reduce", "eliminate → reduce"),
]

CURATED_FLAGSHIP_REWRITES = {
    "Enforce runtime governance for AI at scale. ServiceNow Super Intelligence Control Tower gives you a central hub to discover, secure, govern, observe, and measure any AI across your enterprise.":
        "Enforce runtime governance for AI at scale. Super Intelligence Control Tower gives you a central hub to discover, secure, govern, observe, and measure enterprise AI.",
    "How is Super Intelligence Control Tower different from other AI governance solutions?":
        "How is SI Control Tower unique?",
    "Can Super Intelligence Control Tower help discover AI agents we don't know about?":
        "Can SI Control Tower discover shadow AI agents?",
    "How does Super Intelligence Control Tower handle regulatory compliance?":
        "How does SI Control Tower handle compliance?",
    "Does Super Intelligence Control Tower work with our existing ServiceNow investments?":
        "Does SI Control Tower work with ServiceNow?",
    "Most vendors focus on a single area like governance or security or limit capabilities to AI within their own cloud ecosystems. Super Intelligence Control Tower is the only platform that unifies AI discovery, observability, governance, security, and value calculation, to provide you with a single source of truth for every AI asset across every cloud and vendor.":
        "Most vendors only cover one area or their own cloud. Super Intelligence Control Tower unifies AI discovery, observability, governance, security, and value metrics to give you one source of truth for AI assets across clouds and vendors.",
    "Super Intelligence Control Tower provides you with flexible, automated, auditable, and traceable compliance gates. It has out-of-the-box governance workflows and content packs for EU AI Act, NIST AI RMF, and state regulations such as the California AI Transparency Act.":
        "SI Control Tower provides automated, auditable compliance gates with out-of-the-box governance workflows and content packs for the EU AI Act, NIST AI RMF, and state regulations like the California AI Transparency Act.",
    "Absolutely. Super Intelligence Control Tower is fully integrated into the ServiceNow AI Platform and exists within your CMDB. Reach out to your account team to learn more about the compounding benefits of Super Intelligence Control Tower with your existing ServiceNow investments.":
        "Yes. SI Control Tower is built into the ServiceNow AI Platform and your CMDB. Contact your account team to see how it compounds the value of your existing ServiceNow investments.",
    "Review our product doc to learn more about what you’ll get from ServiceNow Super Intelligence Control Tower.":
        "Review our product doc to learn what you’ll get from SI Control Tower.",
    "ServiceNow Super Intelligence Control Tower dashboard showing top action items, inventory of AI assets, and governance posture.":
        "SI Control Tower dashboard showing top action items, inventory of AI assets, and governance posture.",
    "Video showing Super Intelligence Control Tower dashboard showing inventory metrics, agent cards, MCP server list, and revenue tracking.":
        "SI Control Tower dashboard with inventory metrics, agent cards, MCP server list, and revenue tracking.",
    "Video showing Super Intelligence Control Tower security dashboard with 89% score gauge, access controls, least-privilege enforcement, and threat monitoring tiles.":
        "SI Control Tower security dashboard with 89% score gauge, access controls, and threat monitoring tiles.",
    "Video showing Super Intelligence Control Tower governance view with architecture diagrams, policy enforcement tiles, compliance controls, and more.":
        "SI Control Tower governance view with architecture diagrams, policy enforcement, and compliance controls.",
    "Video showing Super Intelligence Control Tower monitoring dashboard with performance gauges, agent metrics, thresholds, and real-time activity tracking tiles.":
        "SI Control Tower monitoring dashboard with performance gauges, agent metrics, and activity tracking tiles.",
    "Video showing Super Intelligence Control Tower business metrics dashboard showing revenue impact ($726K, $92.2K, $23.0K), adoption rates, and ROI trends.":
        "SI Control Tower business metrics dashboard showing revenue impact, adoption rates, and ROI trends.",
    "ServiceNow Super Intelligence Control Tower gives Chief AI Officers the tools to connect AI strategy, governance, and performance enterprise-wide.":
        "SI Control Tower gives Chief AI Officers tools to connect AI strategy, governance, and performance.",
    "ServiceNow Super Intelligence Control Tower gives CIOs and CTOs the real-time insights they need to prove the ROI of any AI .":
        "SI Control Tower gives CIOs and CTOs real-time insights to prove the ROI of enterprise AI.",
    "ServiceNow Super Intelligence Control Tower enables robust AI governance to proactively manage risks, and safeguard operations and reputation.":
        "SI Control Tower enables powerful AI governance to manage risk and safeguard operations and reputation.",
    "Two ServiceNow practitioners explaining how Super Intelligence Control Tower manages AI as an enterprise capability across every use case.":
        "Two ServiceNow practitioners explaining how SI Control Tower manages AI across enterprise use cases.",
}

def make_si_short_target(cur_text, full_target_text, allow_tagline_short=False):
    """
    Produces the 16-char 'SI Control Tower' short variant (exact same 16-char length as 'AI Control Tower'),
    plus localized SI short variants across all 10 non-English locales.
    Note: Keep company positioning tagline ('the Super Intelligence Control Tower for business reinvention') in full
    unless inside a strict 50-char heading cap (allow_tagline_short=True).
    """
    if "business reinvention" in cur_text.lower() and not allow_tagline_short:
        return full_target_text
    short_t = full_target_text.replace("ServiceNow Super Intelligence Control Tower", "SI Control Tower")
    short_t = short_t.replace("Super Intelligence Control Tower", "SI Control Tower")
    localized_short_swaps = [
        ("Super-Intelligenz-Kontrollturm", "SI-Kontrollturm"),
        ("Superintelligenz-Kontrollturm", "SI-Kontrollturm"),
        ("superintelligenz-kontrollturm", "SI-Kontrollturm"),
        ("Tour de contrôle de la Super Intelligence", "Tour de contrôle SI"),
        ("Tour de contrôle de la superintelligence", "Tour de contrôle SI"),
        ("Tour de contrôle Super Intelligence", "Tour de contrôle SI"),
        ("Tour de contrôle superintelligence", "Tour de contrôle SI"),
        ("tour de contrôle de la Super Intelligence", "tour de contrôle SI"),
        ("tour de contrôle de la superintelligence", "tour de contrôle SI"),
        ("tour de contrôle Super Intelligence", "tour de contrôle SI"),
        ("tour de contrôle superintelligence", "tour de contrôle SI"),
        ("Torre de control de Superinteligencia", "Torre de control de SI"),
        ("Torre de control de superinteligencia", "Torre de control de SI"),
        ("torre de control de Superinteligencia", "torre de control de SI"),
        ("torre de control de superinteligencia", "torre de control de SI"),
        ("Torre de controle de Superinteligência", "Torre de controle de SI"),
        ("Torre de controle de superinteligência", "Torre de controle de SI"),
        ("torre de controle de Superinteligência", "torre de controle de SI"),
        ("torre de controle de superinteligência", "torre de controle de SI"),
        ("Torre di controllo della Super Intelligenza", "Torre di controllo SI"),
        ("Torre di controllo della superintelligenza", "Torre di controllo SI"),
        ("torre di controllo della Super Intelligenza", "torre di controllo SI"),
        ("torre di controllo della superintelligenza", "torre di controllo SI"),
        ("Superintelligentie-controletoren", "SI-controletoren"),
        ("superintelligentie-controletoren", "SI-controletoren"),
        ("controletoren voor Superintelligentie", "controletoren voor SI"),
        ("controletoren voor superintelligentie", "controletoren voor SI"),
        ("スーパーインテリジェンス・コントロールタワー", "SI コントロールタワー"),
        ("スーパーインテリジェンス コントロールタワー", "SI コントロールタワー"),
        ("슈퍼 인텔리전스 컨트롤 타워", "SI 컨트롤 타워"),
        ("슈퍼 인텔리전스 컨트롤타워", "SI 컨트롤타워"),
    ]
    for full_loc, short_loc in localized_short_swaps:
        short_t = short_t.replace(full_loc, short_loc)
    return short_t

def audit_wwsg_instance(inst, page_url=""):
    loc = inst.get("location", "")
    cat = inst.get("category", "")
    tgt = inst.get("target_text", "").strip()
    cur = inst.get("current_text", "").strip()
    is_flagship_page = ("ai-control-tower.html" in page_url)

    limit = None
    comp_name = "Body copy"
    if "arc-secondary-nav" in loc:
        limit = 50
        comp_name = "Secondary nav title"
    elif "arc-marquee-large" in loc and ("heading" in loc.lower() or "Eyebrow" in loc):
        limit = 50
        comp_name = "Marquee large heading"
    elif "arc-marquee-large" in loc and "<p>" in loc:
        limit = 175
        comp_name = "Marquee large description"
    elif "arc-marquee" in loc and ("heading" in loc.lower() or "H1" in loc):
        limit = 50
        comp_name = "Marquee heading"
    elif "arc-marquee" in loc and "<p>" in loc:
        limit = 150
        comp_name = "Marquee description"
    elif "arc-accordion-item" in loc and "heading" in loc:
        limit = 50
        comp_name = "Accordion title"
    elif "arc-accordion-item" in loc and "copy-block" in loc:
        limit = 250
        comp_name = "Accordion body"
    elif "arc-next-steps" in loc and "<p>" in loc:
        limit = 150
        comp_name = "Next steps description"
    elif "action-item" in loc and "<p>" in loc:
        limit = 100
        comp_name = "Next steps action item"
    elif "arc-universal-group" in loc and "<p>" in loc:
        limit = 100
        comp_name = "Card description"
    elif ("Heading" in loc or "H1" in loc or "H2" in loc or "H3" in loc) and "Body Copy" not in loc and len(cur) <= 50:
        limit = 50
        comp_name = "Component heading"
    elif "Breadcrumb" in loc:
        limit = 50
        comp_name = "Breadcrumb item"
    elif cat == "Image Alt / ARIA Attribute" or "alt=" in loc:
        limit = 125
        comp_name = "Image alt text"
    elif "Document <title>" in loc or "og:title" in loc or "twitter:title" in loc:
        limit = 60
        comp_name = "SEO <title>"
    elif "description" in loc and cat == "Head <title> & <meta>":
        limit = 160
        comp_name = "SEO <meta description>"

    si_short = make_si_short_target(cur, tgt, allow_tagline_short=(limit == 50))
    t_len = len(tgt)
    si_len = len(si_short)

    used_si_short = False
    recommended_auto_target = tgt
    if limit and t_len > limit:
        if si_len <= limit:
            used_si_short = True
            recommended_auto_target = si_short
        elif not is_flagship_page:
            used_si_short = True
            recommended_auto_target = si_short

    if limit:
        if t_len <= limit:
            char_check = f"{comp_name}: [{t_len}/{limit} chars ✓ Full Name Fits]"
            is_over = False
        elif si_len <= limit:
            char_check = f"{comp_name}: [Full={t_len}/{limit} ⚠️ → Use 'SI Control Tower'={si_len}/{limit} chars ✓]"
            is_over = False
        elif not is_flagship_page:
            char_check = f"{comp_name}: [Pre-existing >{limit} chars → Auto-swap 'SI Control Tower' ({si_len} chars, +0 net delta) ✓]"
            is_over = False
        else:
            over_by = si_len - limit
            char_check = f"{comp_name}: [{si_len}/{limit} chars ⚠️ OVER by {over_by} — needs manual copy trim]"
            is_over = True
    else:
        char_check = f"{comp_name}: [Full={t_len} chars | Short 'SI'={si_len} chars]"
        is_over = False

    flags = []
    rewrite = recommended_auto_target

    if re.search(r"\b(?:AICT|SICT)\b", cur):
        flags.append('[ABBREVIATION: Replace "AICT" with "SI Control Tower" (16 chars) or "Super Intelligence Control Tower"]')

    if "ServiceNow Super Intelligence Control Tower" in rewrite and cat != "Head <title> & <meta>":
        flags.append('[NAMING: Omit redundant "ServiceNow" before product name (-11 chars)]')
        rewrite = rewrite.replace("ServiceNow Super Intelligence Control Tower", "Super Intelligence Control Tower")

    for pat, rep, label in BUZZWORDS_MAP:
        if re.search(pat, rewrite, flags=re.I):
            flags.append(f"[BUZZWORD: {label}]")
            rewrite = re.sub(pat, rep, rewrite, flags=re.I)

    for pat, rep, label in ABSOLUTE_MAP:
        if re.search(pat, rewrite, flags=re.I):
            flags.append(f"[ABSOLUTE TERM: {label}]")
            rewrite = re.sub(pat, rep, rewrite, flags=re.I)

    if cat == "Image Alt / ARIA Attribute" or "alt=" in loc:
        if re.match(r"^(?:image|photo|screenshot|video)\s+(?:of|showing)\s+", rewrite, flags=re.I):
            flags.append('[ACCESSIBILITY: Strip "Image/Screenshot/Video showing" prefix from alt text]')
            rewrite = re.sub(r"^(?:image|photo|screenshot|video)\s+(?:of|showing)\s+", "", rewrite, flags=re.I)
            if rewrite:
                rewrite = rewrite[0].upper() + rewrite[1:]

    if tgt in CURATED_FLAGSHIP_REWRITES:
        rewrite = CURATED_FLAGSHIP_REWRITES[tgt]
    elif (cat == "Image Alt / ARIA Attribute" or "alt=" in loc) and len(rewrite) > 125:
        rewrite = make_si_short_target(cur, rewrite)
        rewrite = rewrite.replace(" showing ", " with ").replace(", and more.", ".")
        if len(rewrite) > 125:
            cut = rewrite[:121].rsplit(" ", 1)[0].rstrip(",;:-")
            cut = re.sub(r"\s+(?:and|or|with|for|to|in|on|of|across|the|a|an)$", "", cut, flags=re.I)
            rewrite = cut + "."

    has_an_article = bool(re.search(r"\ban\s+AI\s+Control\s+Tower\b", cur, flags=re.I))
    if has_an_article:
        flags.append('[GRAMMAR NOTE: "an AI Control Tower" → use "an SI Control Tower" or "a Super Intelligence Control Tower"]')

    # Determine Execution Mode for this specific instance:
    # Automated Swap vs. Manual Copy Edit (Clean cutover — H1 Hero is 100% Automated!)
    needs_manual_copy = False
    manual_reason = ""
    if is_over:
        needs_manual_copy = True
        manual_reason = f"Exceeds {comp_name} {limit}-char cap — use trimmed WWSG rewrite"
    elif tgt in CURATED_FLAGSHIP_REWRITES and ("Accordion" in comp_name or "Marquee large description" in comp_name):
        needs_manual_copy = True
        manual_reason = f"Flagship {comp_name} rewrite recommended for WWSG length/clarity"

    rw_len = len(rewrite)
    if limit:
        rw_badge = f"[{rw_len}/{limit} chars {'✓' if rw_len <= limit else '⚠️'}]"
    else:
        rw_badge = f"[{rw_len} chars]"

    return {
        "char_check": char_check,
        "is_over": is_over,
        "used_si_short": used_si_short,
        "si_short_target": si_short,
        "recommended_auto_target": recommended_auto_target,
        "flags": flags,
        "wwsg_rewrite": rewrite if (is_over or flags or needs_manual_copy) and rewrite != tgt else None,
        "wwsg_rewrite_badge": rw_badge,
        "needs_manual_copy": needs_manual_copy,
        "manual_reason": manual_reason,
        "comp_name": comp_name,
        "limit": limit,
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
            MASTER_IMAGES_MAP[_im_url] = _im_info
        for _pg_url, _im_list in _m_ocr.get("page_to_images", {}).items():
            _clean_list = [u for u in _im_list if not u.lower().endswith(".pdf") and "logo-aicte.svg" not in u.lower()]
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

total_baked_images_count = sum(1 for v in MASTER_IMAGES_MAP.values() if v.get("has_baked_aict_pixels"))
total_auto_slug_images_count = sum(1 for v in MASTER_IMAGES_MAP.values() if not v.get("has_baked_aict_pixels"))
total_zero_pixel_images_count = total_auto_slug_images_count
total_impacted_images_count = len(MASTER_IMAGES_MAP)

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
    Splits all changes for page `r` into two distinct tracks:
      1. Automated Updates (`auto_html`, `auto_csv`, `auto_edit_count`):
         - <title> & <meta> swaps
         - Automated body copy swaps (where not needs_manual_copy)
         - Automated in-body link slug swaps
         - Automated <arc-image alt> attribute swaps
         - Automated 0-pixel DAM image slug moves
         - Sitewide Chrome XF note
      2. Manual Updates (`manual_html`, `manual_csv`, `manual_task_count`, `manual_reasons`, `manual_filter_tags`):
         - Manual WWSG Copy Edits (where needs_manual_copy == True)
         - Manual Baked-In Image Re-renders (now-ai-control-tower, secure-any-ai-action)
         - Manual Live/Vector PDF Updates
         - Manual L10n Native Grammar Review (10 non-English flagship product pages)
         - URL Slug 301 Redirect Note (cross-linked to Tab 3: 301 Redirects)
    """
    u = full_url(r["url"])
    loc_code = r.get("locale", get_page_locale(u))
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

    auto_html_items = []
    auto_csv_items = []
    manual_html_items = []
    manual_csv_items = []
    wwsg_csv_notes = []

    auto_text_meta_alt_cnt = 0
    auto_link_cnt = len(body_slug_links)
    manual_copy_cnt = 0
    si_short_cnt = 0
    manual_reasons = []
    manual_filter_tags = set()

    # A. URL Slug / 301 Redirect (Manual / Web Ops task, cross-linked to Tab 3)
    if "ai-control-tower" in u.lower() or "aict" in u.lower():
        new_u = re.sub(r'ai-control-tower', 'super-intelligence-control-tower', u, flags=re.I)
        new_u = re.sub(r'aict', 'super-intelligence-control-tower', new_u, flags=re.I)
        manual_reasons.append("AEM Page Move & 301 Redirect")
        manual_filter_tags.add("m-slug")
        manual_html_items.append(
            f"<li><span class='badge b-w1'>🛠️ MANUAL: AEM PAGE MOVE &amp; 301</span> "
            f"<span class='badge b-w3' style='cursor:pointer;margin-left:4px;' onclick=\"showTab('tab-redirects', document.getElementById('btn-tab-redirects'))\">🔀 View in Tab 3 (301 Redirects) ↗</span><br>"
            f"<strong>URL Slug &amp; 301 Redirect:</strong> {a_link(u, None, 'var(--blue)')} &rarr; <code>{html.escape(new_u)}</code> "
            f"(move AEM page node, deploy HTTP 301 redirect rule, and update canonical, <code>og:url</code> &amp; <code>hreflang</code> cluster).</li>"
        )
        manual_csv_items.append(f"[MANUAL: URL SLUG & 301] Move {u} -> {new_u}")

    # B. <title> & <meta> Tag Instances (Automated)
    if meta_instances:
        meta_lines = []
        meta_csv_parts = []
        for idx, m in enumerate(meta_instances, 1):
            m_loc = m.get("location", "meta")
            m_old = m.get("current_text", "")
            m_new = m.get("target_text", "")
            w_res = audit_wwsg_instance(m, u)
            if w_res["used_si_short"]:
                si_short_cnt += 1
            auto_text_meta_alt_cnt += 1
            si_opt = w_res["si_short_target"]
            mode_badge = "<span class='badge b-ok' style='font-size:10px;padding:1px 6px;'>🤖 AUTO SWAP</span>"
            cc_badge = f"<span style='color:{'var(--amber)' if w_res['is_over'] else 'var(--muted)'};font-size:11px;margin-left:6px;'>{html.escape(w_res['char_check'])}</span>"
            seo_note = ""
            if "Document <title>" in m_loc:
                if len(m_new) <= 60:
                    seo_note = f"<br><span style='color:var(--blue);font-size:11px;'>• <strong>SEO/AIO <code>&lt;title&gt;</code> Guidance:</strong> Full name <code>\"{html.escape(m_new)}\"</code> ({len(m_new)}/60 chars) fits Google SERP limit &amp; establishes canonical LLM entity. Short alternative: <code>\"{html.escape(si_opt)}\"</code> ({len(si_opt)}/60 chars).</span>"
                else:
                    seo_note = f"<br><span style='color:var(--amber);font-size:11px;'>• <strong>SEO/AIO <code>&lt;title&gt;</code> Guidance:</strong> Full name is {len(m_new)}/60 chars (truncates in SERP) &rarr; <strong>Auto-swaps to short name:</strong> <code>\"{html.escape(si_opt)}\"</code> ({len(si_opt)}/60 chars ✓).</span>"
            meta_lines.append(
                f"<div style='margin-top:3px;padding-left:6px;border-left:2px solid var(--border);'>"
                f"{mode_badge} <strong>Meta #{idx}</strong> <code>{html.escape(m_loc)}</code> {cc_badge}:<br>"
                f"<span style='color:var(--muted);'>Current:</span> <code>\"{html.escape(m_old)}\"</code><br>"
                f"<span style='color:var(--green);'>Target:</span> <strong>\"{html.escape(w_res['recommended_auto_target'])}\"</strong>"
                f"{seo_note}</div>"
            )
            meta_csv_parts.append(f"Meta #{idx} [AUTO] ({m_loc}) {w_res['char_check']}: \"{m_old}\" -> \"{w_res['recommended_auto_target']}\"")
        auto_html_items.append(
            f"<li><strong>Page <code>&lt;title&gt;</code> &amp; <code>&lt;meta&gt;</code> Tags ({len(meta_instances)} automated tags — Meta #1 to Meta #{len(meta_instances)}):</strong>"
            f"{''.join(meta_lines)}</li>"
        )
        auto_csv_items.append(f"[AUTO: TITLE & META ({len(meta_instances)})] " + " || ".join(meta_csv_parts))

    # C. Visible Body Copy Instances — Split into Automated vs. Manual WWSG Copy Edits
    if body_text_instances:
        total_b = len(body_text_instances)
        auto_body_entries = []
        manual_body_entries = []
        auto_body_csv = []
        manual_body_csv = []

        for idx, b in enumerate(body_text_instances, 1):
            b_idx = b.get("instance_num", idx)
            b_loc = b.get("location", "Body Copy")
            b_term = b.get("matched_term", "")
            b_old = b.get("current_text", "")
            b_new = b.get("target_text", "")
            w_res = audit_wwsg_instance(b, u)
            if w_res["used_si_short"]:
                si_short_cnt += 1
            si_opt = w_res["si_short_target"]

            cc_color = "var(--amber)" if w_res["is_over"] else "var(--green)"
            cc_html = f"<span style='color:{cc_color};font-size:11px;font-weight:600;margin-left:6px;'>• {html.escape(w_res['char_check'])}</span>"
            flags_html = ""
            if w_res["flags"]:
                fl_str = " ".join(w_res["flags"])
                flags_html = f"<br><span style='color:var(--amber);font-size:11px;'><strong>WWSG Notes:</strong> {html.escape(fl_str)}</span>"
                wwsg_csv_notes.append(f"Body #{b_idx}: {w_res['char_check']} {fl_str}")
            elif w_res["is_over"]:
                wwsg_csv_notes.append(f"Body #{b_idx}: {w_res['char_check']}")

            si_line = ""
            if si_opt != b_new and "business reinvention" not in b_old.lower():
                si_line = f"<br><span style='color:var(--blue);'>Short Option (<code>SI Control Tower</code> — 0 char expansion):</span> <code>\"{html.escape(si_opt)}\"</code>"

            rw_html = ""
            if w_res["wwsg_rewrite"]:
                rw_html = (
                    f"<br><span style='color:var(--green);font-weight:600;'>Recommended WWSG Rewrite {html.escape(w_res['wwsg_rewrite_badge'])}:</span> "
                    f"<strong style='color:#93c5fd;'>\"{html.escape(w_res['wwsg_rewrite'])}\"</strong>"
                )

            if w_res["needs_manual_copy"]:
                manual_copy_cnt += 1
                mode_badge = "<span class='badge b-w1' style='font-size:10px;padding:1px 6px;'>🛠️ MANUAL COPY EDIT</span>"
                reason_line = f"<br><span style='color:var(--amber);font-size:11px;'><strong>Why Manual:</strong> {html.escape(w_res['manual_reason'])}</span>"
                item_html = (
                    f"<div style='margin-top:4px;padding-left:6px;border-left:2px solid var(--amber);'>"
                    f"{mode_badge} <strong>Instance #{b_idx}</strong> (of {total_b}) <span style='color:var(--blue);font-weight:600;'>[{html.escape(b_loc)}]</span> "
                    f"<span style='color:var(--muted);font-size:11px;'>(matched <code>{html.escape(b_term)}</code>)</span>{cc_html}:<br>"
                    f"<span style='color:var(--muted);'>Current:</span> <code>\"{html.escape(b_old)}\"</code><br>"
                    f"<span style='color:var(--green);'>Full Target:</span> <strong>\"{html.escape(b_new)}\"</strong>"
                    f"{si_line}{reason_line}{flags_html}{rw_html}</div>"
                )
                manual_body_entries.append(item_html)
                rec_txt = w_res["wwsg_rewrite"] or w_res["recommended_auto_target"]
                manual_body_csv.append(f"#{b_idx} [MANUAL_COPY | {b_loc}] \"{b_old}\" -> \"{rec_txt}\"")
            else:
                auto_text_meta_alt_cnt += 1
                if w_res["used_si_short"]:
                    mode_badge = "<span class='badge b-ok' style='font-size:10px;padding:1px 6px;'>🤖 AUTO SWAP (Uses 'SI Control Tower')</span>"
                else:
                    mode_badge = "<span class='badge b-ok' style='font-size:10px;padding:1px 6px;'>🤖 AUTO SWAP</span>"
                item_html = (
                    f"<div style='margin-top:4px;padding-left:6px;border-left:2px solid var(--blue);'>"
                    f"{mode_badge} <strong>Instance #{b_idx}</strong> (of {total_b}) <span style='color:var(--blue);font-weight:600;'>[{html.escape(b_loc)}]</span> "
                    f"<span style='color:var(--muted);font-size:11px;'>(matched <code>{html.escape(b_term)}</code>)</span>{cc_html}:<br>"
                    f"<span style='color:var(--muted);'>Current:</span> <code>\"{html.escape(b_old)}\"</code><br>"
                    f"<span style='color:var(--green);'>Automated Target:</span> <strong>\"{html.escape(w_res['recommended_auto_target'])}\"</strong>"
                    f"{si_line}{flags_html}{rw_html}</div>"
                )
                auto_body_entries.append(item_html)
                m_tag = "AUTO_SI_SHORT" if w_res["used_si_short"] else "AUTO_FULL"
                auto_body_csv.append(f"#{b_idx} [{m_tag} | {b_loc}] \"{b_old}\" -> \"{w_res['recommended_auto_target']}\"")

        if auto_body_entries:
            is_flagship = ("products/ai-control-tower.html" in u or "contact-us-aict.html" in u)
            if len(auto_body_entries) <= 6 or is_flagship:
                auto_html_items.append(
                    f"<li><strong>Automated Body Copy, Headings, Cards &amp; Breadcrumbs ({len(auto_body_entries)} of {total_b} body instances):</strong>"
                    f"{''.join(auto_body_entries)}</li>"
                )
            else:
                shown = auto_body_entries[:5]
                rest = auto_body_entries[5:]
                extra_html = (
                    f"<details style='margin-top:5px;cursor:pointer;'>"
                    f"<summary style='color:var(--blue);font-weight:600;'>+ Show all remaining {len(rest)} automated body copy instances on this page</summary>"
                    f"<div style='margin-top:4px;'>{''.join(rest)}</div></details>"
                )
                auto_html_items.append(
                    f"<li><strong>Automated Body Copy, Headings, Cards &amp; Breadcrumbs ({len(auto_body_entries)} of {total_b} body instances):</strong>"
                    f"{''.join(shown)}{extra_html}</li>"
                )
            auto_csv_items.append(f"[AUTO: BODY INSTANCES ({len(auto_body_entries)})] " + " || ".join(auto_body_csv))

        if manual_body_entries:
            manual_reasons.append(f"Copy Edit ({manual_copy_cnt} WWSG / FAQ / Marquee items)")
            manual_filter_tags.add("m-copy")
            manual_html_items.append(
                f"<li><strong>Manual WWSG Copy Edits ({len(manual_body_entries)} of {total_b} body instances — pre-existing copy over WWSG char cap):</strong>"
                f"{''.join(manual_body_entries)}</li>"
            )
            manual_csv_items.append(f"[MANUAL: WWSG COPY EDITS ({len(manual_body_entries)})] " + " || ".join(manual_body_csv))

    elif body_ai_ct > 0 or raw_loc > 0 or raw_aict > 0:
        cnt = max(1, body_ai_ct + raw_loc + raw_aict)
        auto_text_meta_alt_cnt += cnt
        auto_html_items.append(
            f"<li><span class='badge b-ok'>🤖 AUTO SWAP</span> <strong>On-Page Body Component / Card Copy ({cnt} instance{'s' if cnt != 1 else ''}):</strong> Update <code>AI Control Tower</code> / <code>AICT</code> &rarr; <strong>Super Intelligence Control Tower</strong> (or <strong>SI Control Tower</strong> on tight UI cards).</li>"
        )
        auto_csv_items.append(f"[AUTO: BODY COMPONENT COPY ({cnt}x)] Update AI Control Tower / AICT -> 'Super Intelligence Control Tower' or 'SI Control Tower'")

    # D. In-Body Anchor Links & Component Data Attributes (Automated)
    if body_slug_links or data_attr_instances or body_slug > 0:
        link_lines = []
        link_csv_parts = []
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
            lk_old_clickable = a_link(full_url(lk_old), lk_old, 'var(--blue)') if (lk_old.startswith("/") or lk_old.startswith("http")) else f"<code>{html.escape(lk_old)}</code>"
            link_lines.append(
                f"<div style='margin-top:3px;padding-left:6px;border-left:2px solid var(--border);'>"
                f"<span class='badge b-ok' style='font-size:10px;padding:1px 6px;'>🤖 AUTO LINK SWAP</span> <strong>Link #{idx}</strong> <code>&lt;{html.escape(lk_tag)} {html.escape(lk_attr)}=...&gt;</code>:<br>"
                f"{lk_old_clickable} &rarr; <code>{html.escape(lk_new)}</code></div>"
            )
            link_csv_parts.append(f"Link #{idx} [AUTO] (<{lk_tag} {lk_attr}>): {lk_old} -> {lk_new}")
        if body_slug > 0:
            link_lines.append(
                f"<div style='margin-top:3px;color:var(--muted);'>• Total in-page <code>ai-control-tower</code> slug references: <strong>{body_slug}</strong> (automated via Groovy link rule).</div>"
            )
            link_csv_parts.append(f"Total in-page slug refs ({body_slug}x) [AUTO]: ai-control-tower -> super-intelligence-control-tower")
        auto_html_items.append(
            f"<li><strong>Automated In-Body CTA Links &amp; Slug References ({len(body_slug_links)} in-body links, {body_slug} total slug refs):</strong>"
            f"{''.join(link_lines)}</li>"
        )
        auto_csv_items.append("[AUTO: LINKS & SLUG REFS] " + " || ".join(link_csv_parts))

    # E. On-Page DAM Images & HTML alt Attributes — Split into Auto (alt swaps + 0-pixel DAM moves) vs. Manual (Baked-In Images)
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

    if auto_imgs or alt_instances:
        auto_img_lines = []
        for idx, im in enumerate(auto_imgs, 1):
            fim = full_url(im)
            auto_img_lines.append(
                f"<div style='margin-top:3px;'>🖼️ <strong>Auto DAM Image #{idx}:</strong> {a_link(fim, None, 'var(--purple)')} "
                f"<span class='badge b-ok'>🤖 AUTO DAM MOVE</span> <strong>[0 Baked Pixels — Rename DAM file slug + 301]</strong></div>"
            )
        alt_html = ""
        alt_csv_parts = []
        if alt_instances:
            alt_items_html = []
            for a_idx, a_obj in enumerate(alt_instances, 1):
                a_loc = a_obj.get("location", "<arc-image alt=\"...\">")
                a_old = a_obj.get("current_text", "")
                w_res = audit_wwsg_instance(a_obj, u)
                if w_res["used_si_short"]:
                    si_short_cnt += 1
                auto_text_meta_alt_cnt += 1
                best_alt = w_res["wwsg_rewrite"] or w_res["recommended_auto_target"]
                mode_badge = "<span class='badge b-ok' style='font-size:10px;padding:1px 6px;'>🤖 AUTO ALT SWAP</span>"
                cc_color = "var(--amber)" if w_res["is_over"] else "var(--green)"
                cc_html = f"<span style='color:{cc_color};font-size:11px;font-weight:600;margin-left:6px;'>• {html.escape(w_res['char_check'])}</span>"
                alt_items_html.append(
                    f"<div style='margin-top:3px;padding-left:6px;border-left:2px solid var(--purple);'>"
                    f"{mode_badge} <strong>Alt #{a_idx}</strong> (<code>{html.escape(a_loc)}</code>){cc_html}:<br>"
                    f"<span style='color:var(--muted);'>Current:</span> <code>alt=\"{html.escape(a_old)}\"</code><br>"
                    f"<span style='color:var(--green);'>Recommended Target (≤125 chars ✓):</span> <strong>alt=\"{html.escape(best_alt)}\"</strong></div>"
                )
                alt_csv_parts.append(f"Alt #{a_idx} [AUTO] ({len(best_alt)}/125 chars): \"{a_old}\" -> \"{best_alt}\"")
            alt_html = (
                f"<div style='margin-top:5px;'><strong>Automated Image <code>alt</code> Attributes ({len(alt_instances)} tags — Alt #1 to Alt #{len(alt_instances)}):</strong>"
                f"{''.join(alt_items_html)}</div>"
            )
        auto_html_items.append(
            f"<li><strong>Automated DAM Image Moves ({len(auto_imgs)}) &amp; HTML <code>alt</code> Attributes ({len(alt_instances)}):</strong>{''.join(auto_img_lines)}{alt_html}</li>"
        )
        auto_csv_items.append(
            f"[AUTO: IMAGES ({len(auto_imgs)}) & ALT INSTANCES ({len(alt_instances)})] "
            f"Auto Imgs: {', '.join(full_url(i) for i in auto_imgs)} | Alts: {' || '.join(alt_csv_parts)}"
        )

    if baked_imgs:
        manual_reasons.append(f"Baked-In Image Re-render ({len(baked_imgs)})")
        manual_filter_tags.add("m-img")
        baked_lines = []
        baked_csv_parts = []
        for idx, fim in enumerate(baked_imgs, 1):
            minfo = MASTER_IMAGES_MAP.get(fim, {})
            ocr_disp = clean_ocr_display(minfo.get("matched_ocr_lines", []))
            dims = f"{minfo.get('width', 0)}×{minfo.get('height', 0)}" if minfo.get("width") else "Image"
            kb = f"{minfo.get('bytes', 0)//1024:,} KB" if minfo.get("bytes") else ""
            spec_str = f" ({dims}{', ' + kb if kb else ''})" if dims != "Image" else ""
            slug_note = " + <strong>301 Rename DAM Slug</strong>" if minfo.get("has_aict_slug") else " (Keep existing generic filename)"
            detail = (
                f"<strong>Vision OCR Match{spec_str}:</strong> <code>\"{html.escape(ocr_disp)}\"</code> baked into image pixels — "
                f"re-render/retouch graphic text to <em>\"Super Intelligence Control Tower\"</em> (or <em>\"SI Control Tower\"</em>){slug_note}."
            )
            baked_lines.append(
                f"<div style='margin-top:3px;padding-left:6px;border-left:2px solid var(--amber);'>"
                f"🖼️ <span class='badge b-w1'>🛠️ MANUAL DESIGN RE-RENDER</span> <strong>Baked-In Image #{idx}:</strong> {a_link(fim, None, 'var(--purple)')}<br>"
                f"<span style='color:var(--muted);font-size:11.5px;'>{detail}</span></div>"
            )
            baked_csv_parts.append(f"{fim} (OCR: \"{ocr_disp}\")")
        manual_html_items.append(
            f"<li><strong>Manual Baked-In Image Re-renders ({len(baked_imgs)} asset{'s' if len(baked_imgs) != 1 else ''}):</strong>{''.join(baked_lines)}</li>"
        )
        manual_csv_items.append(f"[MANUAL: BAKED-IN IMAGES ({len(baked_imgs)})] " + " | ".join(baked_csv_parts))


    # F. Linked Non-FAQ PDFs with Internal Matches (Manual)
    live_pdfs = [p for p in pdfs if pdf_match_lookup.get(p, (0,))[0] > 0 or "ds-ai-control-tower" in p]
    if live_pdfs:
        manual_reasons.append(f"PDF Update ({len(live_pdfs)})")
        manual_filter_tags.add("m-pdf")
        pdf_lines = []
        for idx, p in enumerate(live_pdfs, 1):
            m_info = pdf_match_lookup.get(p)
            if m_info and m_info[0] > 0:
                m_cnt, p_pages, _ = m_info
                pdf_lines.append(
                    f"<div style='margin-top:3px;padding-left:6px;border-left:2px solid var(--amber);'>"
                    f"📄 <span class='badge b-w1'>🛠️ MANUAL PDF UPDATE</span> <strong>PDF #{idx}:</strong> {a_link(p, None, 'var(--amber)')} — "
                    f"<strong>{m_cnt} matches across {p_pages} PDF pages</strong> (re-export source PDF &amp; configure DAM 301 redirect).</div>"
                )
            else:
                pdf_lines.append(
                    f"<div style='margin-top:3px;padding-left:6px;border-left:2px solid var(--amber);'>"
                    f"📄 <span class='badge b-w1'>🛠️ MANUAL PDF UPDATE</span> <strong>PDF #{idx}:</strong> {a_link(p, None, 'var(--amber)')} — "
                    f"Update datasheet copy &amp; filename slug.</div>"
                )
        manual_html_items.append(f"<li><strong>Manual Linked DAM PDF Updates ({len(live_pdfs)}):</strong>{''.join(pdf_lines)}</li>")
        manual_csv_items.append(f"[MANUAL: LINKED PDFS ({len(live_pdfs)})] " + ", ".join(live_pdfs))

    # G. Localized Flagship Native Grammar Review (Manual L10n Task on 10 non-English flagship pages)
    if loc_code in ("de", "fr", "fr-ca", "es", "latam", "br", "it", "nl", "jp", "kr") and "products/ai-control-tower.html" in u:
        manual_reasons.append("L10n Native Grammar Review")
        manual_filter_tags.add("m-l10n")
        manual_html_items.append(
            f"<li><span class='badge b-w2'>🛠️ MANUAL L10N REVIEW</span> <strong>Localized Flagship Native Grammar Check (<code>{loc_code.upper()}</code>):</strong> "
            f"Verify gender/article agreement and compound noun inflections on translated flagship headings and FAQs after automated AEM script run.</li>"
        )
        manual_csv_items.append(f"[MANUAL: L10N GRAMMAR REVIEW ({loc_code.upper()})] Verify gender/article agreement on translated flagship page")

    # H. Sitewide Chrome Note (Automated via XF)
    if ch > 0:
        auto_html_items.append(
            f"<li style='color:var(--muted);'><span class='badge b-ok' style='font-size:10px;padding:1px 6px;'>🤖 AUTO VIA XF</span> "
            f"<strong>Sitewide Chrome ({ch} matches):</strong> Global Navigation Mega-Menu XF + Header Search Prompt JSON (auto-updated in Wave 1).</li>"
        )

    auto_edit_count = auto_text_meta_alt_cnt + auto_link_cnt + len(auto_imgs)
    if auto_edit_count == 0:
        auto_edit_count = max(1, r["true_body"] + r.get("raw_localized", 0))

    manual_task_count = (
        (1 if "m-slug" in manual_filter_tags else 0)
        + manual_copy_cnt
        + len(baked_imgs)
        + len(live_pdfs)
        + (1 if "m-l10n" in manual_filter_tags else 0)
    )

    auto_html = f"<ul style='margin:4px 0 0 16px;padding:0;font-size:12px;line-height:1.45;'>{''.join(auto_html_items)}</ul>" if auto_html_items else "<div style='color:var(--muted);font-size:12px;'>No page-level automated items</div>"
    manual_html = f"<ul style='margin:4px 0 0 16px;padding:0;font-size:12px;line-height:1.45;'>{''.join(manual_html_items)}</ul>" if manual_html_items else "<div style='color:var(--muted);font-size:12px;'>0 manual tasks required</div>"

    auto_csv_str = " ; ".join(auto_csv_items) if auto_csv_items else "None"
    manual_csv_str = " ; ".join(manual_csv_items) if manual_csv_items else "None (100% Automated)"
    combined_csv_str = " ; ".join(auto_csv_items + manual_csv_items)
    wwsg_summary_str = " | ".join(wwsg_csv_notes) if wwsg_csv_notes else "Fits WWSG 2026 Q3 limits (via Full Name or 'SI Control Tower')"

    return {
        "is_fully_auto": (manual_task_count == 0),
        "auto_html": auto_html,
        "auto_csv": auto_csv_str,
        "auto_edit_count": auto_edit_count,
        "auto_text_meta_alt_cnt": auto_text_meta_alt_cnt,
        "auto_link_cnt": auto_link_cnt,
        "auto_img_cnt": len(auto_imgs),
        "si_short_cnt": si_short_cnt,
        "manual_html": manual_html,
        "manual_csv": manual_csv_str,
        "manual_task_count": manual_task_count,
        "manual_copy_cnt": manual_copy_cnt,
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
        wwsg_notes = "Global Nav label 'Super Intelligence Control Tower' (32 chars) or 'SI Control Tower' (16 chars)"
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
            "WWSG_Copy_And_Char_Audit": "[NAMING: Use 'Super Intelligence Control Tower' or 'SI Control Tower' in PDF metadata]",
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
    hosts_full = ", ".join(full_url(u) for u, _ in host_list[:4]) if host_list else "Brand 3.0 DAM Inventory"
    alts = []
    for _, a_list in host_list:
        for a in a_list:
            if a not in alts:
                alts.append(a)
    alt_str = " | ".join(alts[:2]) if alts else "(0 alt text matches)"
    dims = f"{minfo.get('width', 0)}×{minfo.get('height', 0)}" if minfo.get("width") else "Image"
    kb = f"{minfo.get('bytes', 0)//1024} KB" if minfo.get("bytes") else ""
    asset_type_str = f"DAM Image ({dims}{', ' + kb if kb else ''})" if dims != "Image" else "DAM Image"
    if is_false_pos:
        note = "[FALSE POSITIVE — DO NOT RENAME] All India Council for Technical Education (AICTE) partner logo"
        prio = "False Positive - Exclude"
        wave_lbl = "Excluded (False Positive)"
        exec_m = "Exclude (False Positive)"
        auto_c, man_c = 0, 0
        auto_bd, man_bd = "None", "None"
    elif is_baked:
        ocr_disp = clean_ocr_display(minfo.get("matched_ocr_lines", []))
        slug_act = " + 301 Rename DAM Slug" if minfo.get("has_aict_slug") else ""
        note = f"[MANUAL: BAKED-IN PIXELS — RE-RENDER GRAPHIC{slug_act}] Vision OCR confirmed baked-in text: \"{ocr_disp}\" | Used on {len(host_list)} pages ({hosts_full}) | Alt: {alt_str}"
        prio = "P0 - Baked-In Image Pixels"
        wave_lbl = "Wave 1"
        exec_m = f"Manual Update (Design Re-render{slug_act})"
        auto_c, man_c = 0, 1
        auto_bd, man_bd = "None", note
    else:
        note = f"[AUTO: DAM IMAGE RENAME + ALT TEXT] 0 baked-in pixels; rename file slug ai-control-tower -> super-intelligence-control-tower & update alt text | Used on {len(host_list)} pages ({hosts_full})"
        prio = "P1 - DAM Image Slug + Alt Text"
        wave_lbl = "Wave 2"
        exec_m = "100% Automated (DAM Move + Alt Swap)"
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
        "WWSG_Copy_And_Char_Audit": "[ACCESSIBILITY: Updated alt text <= 125 chars using 'SI Control Tower' or 'Super Intelligence Control Tower']",
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
# 7. Build Interactive HTML Table Rows for Tab 1 (Automated Updates) & Tab 2 (Manual Updates)
# ---------------------------------------------------------------------------
auto_table_rows_html = []
manual_table_rows_html = []

manual_wave_counts = {"w1": 0, "w2": 0, "w3": 0}
manual_type_counts = {"m-copy": 0, "m-slug": 0, "m-img": 0, "m-pdf": 0, "m-l10n": 0}

for r in unique_body_pages:
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

    loc_badge = f'<span class="badge {loc_cls}" style="margin-left:4px;">{loc_lbl}</span>'
    cat_badge = f'<span class="badge b-warn" style="margin-top:4px;display:inline-block;">{html.escape(an["category"])}</span>'

    visits_val = an["ytd_visits"]
    conv_val = an["ytd_conversions"]
    fcr_val = an["ytd_fcr"] * 100
    rank_val = an["ytd_rank"]
    p_score = an["prio_score"]
    sv_val = an["strategic_visibility"]
    urg_val = an["urgency"]
    en_base_v = an.get("en_baseline_visits", 0)
    en_base_c = an.get("en_baseline_conversions", 0)
    en_base_p = an.get("en_baseline_prio_score", p_score)

    if visits_val > 0:
        analytics_html = (
            f"<div style='font-size:14px;font-weight:700;color:var(--green);'>{visits_val:,} <span style='font-size:11px;font-weight:400;color:var(--muted);'>US YTD Visits</span></div>"
            f"<div style='font-size:12.5px;font-weight:600;color:var(--amber);margin-top:2px;'>{conv_val:,} <span style='font-size:11px;font-weight:400;color:var(--muted);'>Form Conv ({fcr_val:.1f}% FCR)</span></div>"
            f"<div style='font-size:11px;color:var(--muted);margin-top:4px;'>"
            f"• <strong>Priority Score:</strong> <strong style='color:var(--text);'>{p_score:.2f}</strong> (Vis={sv_val}, Urg={urg_val})<br>"
            f"{f'• <strong>Site Traffic Rank:</strong> #{rank_val:,}' if rank_val else ''}"
            f"</div>"
        )
    elif loc_code != "us":
        parent_vis_line = f"• <strong>US Parent Visits:</strong> <strong style='color:var(--green);'>{en_base_v:,}</strong> ({en_base_c:,} conv)<br>" if en_base_v > 0 else ""
        analytics_html = (
            f"<div style='font-size:12.5px;font-weight:700;color:var(--blue);'>{loc_lbl} Regional Counterpart</div>"
            f"<div style='font-size:11px;color:var(--muted);margin-top:3px;line-height:1.45;'>"
            f"• <strong>Included by Content Need:</strong> Regional traffic is not in US workbook; included because page has live AICT instances.<br>"
            f"{parent_vis_line}"
            f"• <strong>Inherited Wave Score:</strong> {en_base_p:.2f} (Vis={sv_val}, Urg={urg_val})"
            f"</div>"
        )
    else:
        analytics_html = (
            f"<div style='font-size:12px;color:var(--muted);'>0 Tracked US Workbook Visits</div>"
            f"<div style='font-size:11px;color:var(--muted);margin-top:3px;'>"
            f"• <strong>Included by Content Need</strong> (Live DOM matches)<br>"
            f"• <strong>Priority Score:</strong> {p_score:.2f} (Vis={sv_val}, Urg={urg_val})"
            f"</div>"
        )

    # --- TAB 1 ROW: AUTOMATED UPDATES ---
    if sp["is_fully_auto"]:
        auto_mode_badge = '<span class="badge b-ok" style="margin-top:4px;display:inline-block;">🤖 100% Automated Page</span>'
        auto_cross_badge = '<div style="margin-top:4px;font-size:11px;color:var(--muted);">✓ 0 Manual Tasks Needed</div>'
        auto_exec_attr = "pure-auto"
    else:
        auto_mode_badge = '<span class="badge b-ok" style="margin-top:4px;display:inline-block;">🤖 Automated Bulk Track</span>'
        auto_cross_badge = (
            f'<div style="margin-top:5px;">'
            f'<span class="badge b-w2" style="cursor:pointer;" onclick="showTab(\'tab-manual\', document.getElementById(\'btn-tab-manual\'))" '
            f'title="Click to inspect the Manual Updates for this URL in Tab 2">'
            f'🛠️ +{sp["manual_task_count"]} Manual Task{"s" if sp["manual_task_count"] != 1 else ""} in Tab 2 ↗</span></div>'
        )
        auto_exec_attr = "hybrid-auto"

    auto_summary_lines = []
    if sp["auto_text_meta_alt_cnt"] > 0:
        auto_summary_lines.append(f"• <strong style='color:var(--green);'>🤖 {sp['auto_text_meta_alt_cnt']}</strong> Auto Text/Meta/Alt Swaps")
    if sp["si_short_cnt"] > 0:
        auto_summary_lines.append(f"• <strong style='color:var(--blue);'>✨ {sp['si_short_cnt']}</strong> Fit via <code>SI Control Tower</code>")
    if sp["auto_link_cnt"] > 0:
        auto_summary_lines.append(f"• <strong>🔗 {sp['auto_link_cnt']}</strong> Auto Link Slug Swaps")
    if sp["auto_img_cnt"] > 0:
        auto_summary_lines.append(f"• <strong>🖼️ {sp['auto_img_cnt']}</strong> Auto 0-Pixel DAM Moves")

    auto_table_rows_html.append(f"""
    <tr data-cat="{wc}" data-b30="{fgrp}" data-exec="{auto_exec_attr}" data-loc="{loc_code}">
      <td data-label="Wave, Region &amp; Mode">{wave_badge}{loc_badge}<br>{auto_mode_badge}{auto_cross_badge}<br>{cat_badge}</td>
      <td data-label="Full Page URL" class="mono">{a_link(u)}</td>
      <td data-label="Traffic &amp; Inclusion">{analytics_html}</td>
      <td data-label="Automated Instances (Groovy Script)">
        <div style="font-size:13.5px;margin-bottom:4px;"><strong>{title_str}</strong></div>
        {sp["auto_html"]}
      </td>
      <td data-label="Automated Handoff Summary" style="font-size:12px;line-height:1.5;">
        <strong style="color:var(--green);font-size:14px;">🤖 {sp["auto_edit_count"]}</strong> Auto Edits<br>
        {"<br>".join(auto_summary_lines)}
        {auto_cross_badge}
      </td>
    </tr>""")

    # --- TAB 2 ROW: MANUAL UPDATES (Only for pages with manual_task_count > 0, including Hybrid URLs!) ---
    if not sp["is_fully_auto"]:
        manual_wave_counts[wc] = manual_wave_counts.get(wc, 0) + 1
        for mtag in sp["manual_filter_tags"]:
            manual_type_counts[mtag] = manual_type_counts.get(mtag, 0) + 1
        mtags_attr = " ".join(sp["manual_filter_tags"])

        man_cross_badge = (
            f'<div style="margin-top:5px;">'
            f'<span class="badge b-ok" style="cursor:pointer;" onclick="showTab(\'tab-automated\', document.getElementById(\'btn-tab-automated\'))" '
            f'title="Click to inspect the Automated Bulk Swaps for this same URL in Tab 1">'
            f'🤖 +{sp["auto_edit_count"]} Auto Swaps in Tab 1 ↗</span></div>'
        )
        man_summary_lines = [f"• <strong style='color:var(--amber);'>🛠️ {html.escape(mr)}</strong>" for mr in sp["manual_reasons"]]

        manual_table_rows_html.append(f"""
    <tr data-cat="{wc}" data-b30="{fgrp}" data-mtypes="{mtags_attr}" data-loc="{loc_code}">
      <td data-label="Wave, Region &amp; Mode">{wave_badge}{loc_badge}<br><span class="badge b-w1" style="margin-top:4px;display:inline-block;">🛠️ Manual Update Queue</span>{man_cross_badge}<br>{cat_badge}</td>
      <td data-label="Full Page URL" class="mono">{a_link(u)}</td>
      <td data-label="Traffic &amp; Inclusion">{analytics_html}</td>
      <td data-label="Manual Tasks &amp; WWSG Copy Edits">
        <div style="font-size:13.5px;margin-bottom:4px;"><strong>{title_str}</strong></div>
        {sp["manual_html"]}
      </td>
      <td data-label="Manual Handoff Summary" style="font-size:12px;line-height:1.5;">
        <strong style="color:var(--amber);font-size:14px;">🛠️ {sp["manual_task_count"]}</strong> Manual Task{"s" if sp["manual_task_count"] != 1 else ""}<br>
        {"<br>".join(man_summary_lines)}
        {man_cross_badge}
      </td>
    </tr>""")

# Add Global Nav Experience Fragment row to Tab 1 (Automated Updates)
auto_table_rows_html.append(f"""
    <tr data-cat="w1" data-b30="cat-platform" data-exec="pure-auto" data-loc="us">
      <td data-label="Wave, Region &amp; Mode"><span class="badge b-w1">Wave 1</span> <span class="badge b-ok">🌐 GLOBAL</span><br><span class="badge b-ok" style="margin-top:4px;display:inline-block;">🤖 100% Automated XF</span><br><span class="badge b-warn" style="margin-top:4px;display:inline-block;">Sitewide Global Nav</span></td>
      <td data-label="Full Page URL" class="mono">{len(ssi_only_pages):,} Standard Pages<br><span style="color:var(--muted);font-size:11px;">(All listed in CSV; sample: {a_link('https://www.servicenow.com/products/it-service-management.html', None, 'var(--blue)')})</span></td>
      <td data-label="Traffic &amp; Inclusion">
        <div style="font-size:14px;font-weight:700;color:var(--green);">6,353,811 <span style="font-size:11px;font-weight:400;color:var(--muted);">Remaining Site Visits</span></div>
        <div style="font-size:11px;color:var(--muted);margin-top:4px;">Updated automatically via 1 Global Nav Experience Fragment publish!</div>
      </td>
      <td data-label="Automated Instances (Groovy Script)">
        <strong>Global Navigation Mega-Menu XF (12–16x) + Header Search Prompt JSON (2x)</strong>
        <ul style="margin:4px 0 0 16px;padding:0;font-size:12px;line-height:1.45;">
          <li><span class="badge b-ok">🤖 100% AUTOMATED</span> <strong>0 page-level edits needed across these {len(ssi_only_pages):,} pages.</strong></li>
          <li><strong>1. Global Navigation Experience Fragment:</strong> Update menu label <code>"AI Control Tower"</code> &rarr; <strong>"Super Intelligence Control Tower"</strong> (or <strong>"SI Control Tower"</strong> if nav width is constrained) and link <code>href</code> &rarr; <code>/products/super-intelligence-control-tower.html</code> across all regional nav fragments.</li>
          <li><strong>2. Header Search Prompt JSON:</strong> Update <code>"How does AI Control Tower work?"</code> &rarr; <strong>"How does SI Control Tower work?"</strong> (or <strong>"How does Super Intelligence Control Tower work?"</strong>).</li>
        </ul>
      </td>
      <td data-label="Automated Handoff Summary"><strong style="color:var(--green);">🤖 2 XF Edits</strong><br><span style="color:var(--muted);font-size:11px;">Auto-cascades across {len(ssi_only_pages):,} pages</span></td>
    </tr>
""")

# ---------------------------------------------------------------------------
# 8. Build Dedicated 301 Redirects Table Rows (Tab 3: Webpage Slugs + DAM PDFs + DAM Images)
# ---------------------------------------------------------------------------
redirect_rows_html = []
redirect_idx = 0
webpage_redirect_cnt = 0
pdf_redirect_cnt = 0
img_redirect_cnt = 0

# 8a. Webpage URL Slug Redirects (34 crawled body-match slug pages + regional live event & webinar slug endpoints)
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
    loc_lbl, loc_cls = LOCALE_BADGE_MAP.get(loc_code, (loc_code.upper(), "b-w3"))
    wc = r["wave_code"]
    w_badge = '<span class="badge b-w1">Wave 1</span>' if wc == "w1" else ('<span class="badge b-w2">Wave 2</span>' if wc == "w2" else '<span class="badge b-w3">Wave 3</span>')

    redirect_rows_html.append(f"""
    <tr data-rtype="webpage" data-rloc="{loc_code}">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Asset Type &amp; Wave"><span class="badge b-ok">🌐 Webpage Slug</span><br><div style="margin-top:4px;">{w_badge} <span class="badge {loc_cls}">{loc_lbl}</span></div></td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_u, None, 'var(--blue)')}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_u)}</code></td>
      <td data-label="HTTP Status &amp; Technical Action">
        <span class="badge b-w1">HTTP 301 Permanent</span><br>
        <span style="color:var(--muted);font-size:11.5px;">Move AEM page node, update self-referencing <code>&lt;link rel="canonical"&gt;</code>, <code>og:url</code>, and 14-locale <code>hreflang</code> cluster.</span>
      </td>
    </tr>""")

# Additional linked regional / webinar / pricing webpage slugs found in in-body links & analytics
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
    loc_lbl, loc_cls = LOCALE_BADGE_MAP.get(loc_code, (loc_code.upper(), "b-w3"))
    w_badge = '<span class="badge b-w1">Wave 1</span>' if wc == "w1" else '<span class="badge b-w2">Wave 2</span>'
    redirect_rows_html.append(f"""
    <tr data-rtype="webpage" data-rloc="{loc_code}">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Asset Type &amp; Wave"><span class="badge b-ok">🌐 Webpage Slug</span><br><div style="margin-top:4px;">{w_badge} <span class="badge {loc_cls}">{loc_lbl}</span></div></td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_u, None, 'var(--blue)')}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_u)}</code></td>
      <td data-label="HTTP Status &amp; Technical Action">
        <span class="badge b-w1">HTTP 301 Permanent</span><br>
        <span style="color:var(--muted);font-size:11.5px;">{html.escape(desc_note)} — redirect legacy CTA links &amp; update AEM node path.</span>
      </td>
    </tr>""")

# 8b. DAM PDF Filename Slug Redirects
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
    loc_lbl, loc_cls = LOCALE_BADGE_MAP.get(p_loc, ("🇺🇸 US", "b-ok"))
    w_badge = '<span class="badge b-w1">Wave 1</span>' if ("ebk-aict-discovery-playbook" in old_pdf or "subscription-unit" in old_pdf or "ds-ai-control-tower" in old_pdf) else '<span class="badge b-w2">Wave 2</span>'
    redirect_rows_html.append(f"""
    <tr data-rtype="pdf" data-rloc="{p_loc}">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Asset Type &amp; Wave"><span class="badge b-w2">📄 DAM PDF Slug</span><br><div style="margin-top:4px;">{w_badge} <span class="badge {loc_cls}">{loc_lbl}</span></div></td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_pdf, None, 'var(--amber)')}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_pdf)}</code></td>
      <td data-label="HTTP Status &amp; Technical Action">
        <span class="badge b-w1">HTTP 301 Permanent</span><br>
        <span style="color:var(--muted);font-size:11.5px;">Publish updated PDF at new DAM path and configure 301 redirect for both <code>.pdf</code> and <code>.pdf.coredownload</code> URLs.</span>
      </td>
    </tr>""")

# 8c. DAM Image Filename Slug Redirects
for im in sorted(img_map.keys()):
    if "logo-aicte.svg" in im.lower():
        continue
    if "ai-control-tower" not in im.lower() and "aict" not in im.lower():
        continue
    old_im = full_url(im)
    if old_im in seen_redirect_urls:
        continue
    seen_redirect_urls.add(old_im)
    redirect_idx += 1
    img_redirect_cnt += 1
    new_im = re.sub(r"ai-control-tower", "super-intelligence-control-tower", old_im, flags=re.I)
    new_im = re.sub(r"aict", "super-intelligence-control-tower", new_im, flags=re.I)
    minfo = MASTER_IMAGES_MAP.get(old_im, {})
    is_baked = minfo.get("has_baked_aict_pixels", False) or "now-ai-control-tower" in old_im or "secure-any-ai-action" in old_im
    w_badge = '<span class="badge b-w1">Wave 1</span>' if is_baked else '<span class="badge b-w2">Wave 2</span>'
    mode_note = "🛠️ Manual Design Re-render + DAM 301" if is_baked else "🤖 Automated 0-Pixel DAM Move + 301"
    redirect_rows_html.append(f"""
    <tr data-rtype="image" data-rloc="us">
      <td data-label="#"><strong>#{redirect_idx}</strong></td>
      <td data-label="Asset Type &amp; Wave"><span class="badge b-warn">🖼️ DAM Image Slug</span><br><div style="margin-top:4px;">{w_badge} <span class="badge b-ok">🌐 GLOBAL</span></div></td>
      <td data-label="Old Live URL (301 Source)" class="mono">{a_link(old_im, None, 'var(--purple)')}</td>
      <td data-label="New Target URL (301 Destination)" class="mono"><code>{html.escape(new_im)}</code></td>
      <td data-label="HTTP Status &amp; Technical Action">
        <span class="badge b-w1">HTTP 301 Permanent</span><br>
        <span style="color:var(--muted);font-size:11.5px;">{mode_note} — prevents broken hotlinks or cached social/OG cards.</span>
      </td>
    </tr>""")

total_redirects_count = redirect_idx

# ---------------------------------------------------------------------------
# 9. Build Category Summary Rows & Media Asset Rows (Tab 5)
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
      <td data-label="Rebrand Pages"><strong>{st["pages"]} total pages</strong><br><span style="color:var(--muted);font-size:11px;">(🇺🇸 {st["us_pages"]} US + 🌍 {st["reg_pages"]} Regional)</span></td>
      <td data-label="Auto vs. Manual"><strong style="color:var(--green);">🤖 {st["auto_pages"]} Pure Auto</strong> / <strong style="color:var(--amber);">🛠️ {st["manual_pages"]} Manual Hybrid</strong><br><span style="color:var(--muted);font-size:11px;">({st["instances"]:,} total DOM edits)</span></td>
      <td data-label="US YTD Visits"><strong style="color:var(--green);">{st["visits"]:,}</strong> US YTD visits<br><span style="color:var(--muted);font-size:11px;">({pct_of_rebrand_visits:.1f}% of rebrand traffic | {pct_of_cat_visits:.1f}% of category)</span></td>
      <td data-label="US Conversions"><strong style="color:var(--amber);">{st["conv"]:,}</strong> conversions<br><span style="color:var(--muted);font-size:11px;">(of {site_conv:,} category conversions)</span></td>
      <td data-label="Wave Split">{" ".join(wave_dist)}</td>
    </tr>""")

media_rows_html = []

# Group the 35 baked-in pixel images into their 19 visual design pairs
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

# Sort visual design groups so welcome-ai-control-tower and flagship product images appear at the top, then by page count descending
def _design_sort_key(item):
    bkey, rlist = item
    max_pages = max((x.get("page_count", 0) for x in rlist), default=0)
    is_welcome = 1 if "welcome-ai-control-tower" in bkey else 0
    is_flagship = 1 if ("now-ai-control-tower" in bkey or "secure-any-ai-action" in bkey) else 0
    return (-is_welcome, -is_flagship, -max_pages, bkey)

for d_idx, (bkey, rlist) in enumerate(sorted(baked_design_groups.items(), key=_design_sort_key), 1):
    rlist.sort(key=lambda x: x["url"])
    primary = rlist[0]
    url_links_html = "<br>".join(a_link(x["url"], None, "var(--purple)") for x in rlist)
    all_hosts = []
    for x in rlist:
        for hp in x.get("host_pages", []):
            if hp not in all_hosts:
                all_hosts.append(hp)
    if all_hosts:
        sample_hosts = ", ".join(a_link(hp, None, "var(--blue)") for hp in all_hosts[:3])
        more_hosts = f" (+ {len(all_hosts)-3} regional/other pages)" if len(all_hosts) > 3 else ""
        hosts_html = f"<div style='margin-top:6px;color:var(--muted);font-size:11px;'><strong>Used on {len(all_hosts)} page{'s' if len(all_hosts)!=1 else ''}:</strong> {sample_hosts}{more_hosts}</div>"
    else:
        hosts_html = "<div style='margin-top:6px;color:var(--muted);font-size:11px;'><strong>Source:</strong> Brand 3.0 Excel DAM Inventory (Unlinked on crawled pages — update in DAM if still active)</div>"

    specs_parts = []
    for x in rlist:
        ext = os.path.splitext(x["url"])[1].upper().lstrip(".") or "IMG"
        specs_parts.append(f"{x.get('width', 0)}×{x.get('height', 0)} {ext} ({x.get('bytes', 0)//1024:,} KB)")
    specs_html = "<br>".join(specs_parts)

    ocr_disp = clean_ocr_display(primary.get("matched_ocr_lines", []))
    has_slug = any(x.get("has_aict_slug") for x in rlist)
    slug_badge = '<span class="badge b-w2" style="margin-top:3px;display:inline-block;">+ 301 DAM Slug</span>' if has_slug else '<span class="badge b-w3" style="margin-top:3px;display:inline-block;">Generic Filename (0 Slug Change)</span>'
    action_tail = "and rename DAM filename slug with HTTP 301 redirect." if has_slug else "and republish in-place (filename has no AICT slug)."

    media_rows_html.append(f"""
<tr>
  <td data-label="Execution Mode"><span class="badge b-w1">🛠️ MANUAL DESIGN #{d_idx}</span><br><div style="margin-top:4px"><span class="badge b-warn">{len(rlist)} Rendition{'s' if len(rlist)!=1 else ''}</span><br>{slug_badge}</div></td>
  <td data-label="Full Clickable Asset URL" class="mono">
    {url_links_html}
    {hosts_html}
  </td>
  <td data-label="Specs">{specs_html}</td>
  <td data-label="OCR / PDFKit Finding"><strong>Apple Vision OCR Match:</strong><br><code>"{html.escape(ocr_disp)}"</code></td>
  <td data-label="Handoff Action"><strong>🛠️ Manual Design / UI Retouch:</strong> Re-render or retouch baked-in <code>AI Control Tower</code> pixel text to <em>"Super Intelligence Control Tower"</em> (or <em>"SI Control Tower"</em>) {action_tail}</td>
</tr>""")

# Add consolidated row for all 28 Zero-Pixel Slug-Matched Images
zero_pixel_links_html = "<br>".join(a_link(x["url"], None, "var(--purple)") for x in zero_pixel_slug_imgs[:12])
if len(zero_pixel_slug_imgs) > 12:
    zero_pixel_links_html += f"<div style='margin-top:4px;color:var(--muted);font-size:11px;'>+ {len(zero_pixel_slug_imgs)-12} additional `.sm`/`.xs` renditions (all {len(zero_pixel_slug_imgs)} listed in Tab 3 &amp; CSV)</div>"

media_rows_html.append(f"""
<tr>
  <td data-label="Execution Mode"><span class="badge b-ok">🤖 AUTO: 0 BAKED PIXELS</span><br><div style="margin-top:4px"><span class="badge b-w2">{len(zero_pixel_slug_imgs)} Slug Images</span></div></td>
  <td data-label="Full Clickable Asset URL" class="mono">
    {zero_pixel_links_html}
  </td>
  <td data-label="Specs">Various PNG / JPG / WebP<br>({len(zero_pixel_slug_imgs)} files)</td>
  <td data-label="OCR / PDFKit Finding"><strong>Apple Vision OCR Verified Clean:</strong><br>0 <code>"AI Control Tower"</code> pixels in any of these {len(zero_pixel_slug_imgs)} images!</td>
  <td data-label="Handoff Action"><strong>🤖 Automated DAM Move + 301 Redirect:</strong> Zero graphic design work needed! Script renames DAM file slugs (<code>ai-control-tower</code> &rarr; <code>super-intelligence-control-tower</code>) and updates HTML <code>alt</code> attributes.</td>
</tr>
<tr>
  <td data-label="Execution Mode"><span class="badge b-w1">🛠️ MANUAL: VECTOR PDF RE-EXPORT</span><br><div style="margin-top:4px"><span class="badge b-warn">6 OCR Matches</span></div></td>
  <td data-label="Full Clickable Asset URL" class="mono">
    {a_link('https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/resource-center/e-book/ebk-aict-discovery-playbook.pdf')}<br>
    <div style="margin-top:6px;color:var(--muted);font-size:11px;"><strong>Linked from:</strong> {a_link('https://www.servicenow.com/standard/resource-center/ebook/ebk-aict-discovery-playbook.html', None, 'var(--blue)')}</div>
  </td>
  <td data-label="Specs">11 Pages, PDF<br>(231 MB)</td>
  <td data-label="OCR / PDFKit Finding"><strong>Vision OCR Caught 6 Matches:</strong> Meta title <code>AICT-DISCOVERY-PLAYBOOK-02</code> + 5 flattened vector text matches on Pages 5, 8, 9.</td>
  <td data-label="Handoff Action"><strong>🛠️ Manual Design Action:</strong> Re-export InDesign source with <em>Super Intelligence Control Tower</em> / <em>SI Control Tower</em> and update metadata title.</td>
</tr>
<tr>
  <td data-label="Execution Mode"><span class="badge b-w1">🛠️ MANUAL: LEGAL PDF UPDATE</span><br><div style="margin-top:4px"><span class="badge b-warn">7 PDFKit Matches</span></div></td>
  <td data-label="Full Clickable Asset URL" class="mono">
    {a_link('https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/legal/ai-control-tower-servicenow-subscription-unit-overview.pdf')}<br>
    <div style="margin-top:6px;color:var(--muted);font-size:11px;"><strong>Linked from:</strong> {a_link('https://www.servicenow.com/products/entitlements-packages.html', None, 'var(--blue)')}</div>
  </td>
  <td data-label="Specs">4 Pages, PDF<br>(109.7 KB)</td>
  <td data-label="OCR / PDFKit Finding"><strong>7 Matches:</strong> Page 1 Title/TOC + Page 3 Section Header &amp; 4 Legal clauses.</td>
  <td data-label="Handoff Action"><strong>🛠️ Manual Legal Action:</strong> Coordinate with Legal to publish updated Subscription Unit Overview PDF and add 301 redirect.</td>
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
    host_links = "<br>".join("• " + a_link(h, None, "var(--blue)") for h in hosts[:3]) if hosts else "• " + a_link("https://www.servicenow.com/standard/resource-center.html", None, "var(--blue)")
    sn = html.escape(" | ".join(p.get("snippets", [])[:2]))
    mt = html.escape(p.get("meta_title", "") or "(None)")
    media_rows_html.append(f"""
<tr>
  <td data-label="Execution Mode"><span class="badge b-w2">🛠️ MANUAL: LIVE DAM PDF</span><br><div style="margin-top:4px"><span class="badge b-warn">{m_cnt} Matches</span></div></td>
  <td data-label="Full Clickable Asset URL" class="mono">{a_link(purl)}<div style="margin-top:6px;color:var(--muted);font-size:11px;"><strong>Linked from:</strong><br>{host_links}</div></td>
  <td data-label="Specs">{p.get('pages', 0)} Pages, PDF<br>({p.get('bytes', 0)//1024:,} KB)</td>
  <td data-label="OCR / PDFKit Finding"><strong>Meta Title:</strong> <code>{mt}</code><br><strong>Extracted Text ({m_cnt} matches):</strong> <code>{sn}</code></td>
  <td data-label="Handoff Action"><strong>🛠️ Manual Doc Team Action:</strong> Update source PDF text &amp; metadata title to <em>Super Intelligence Control Tower</em> / <em>SI Control Tower</em> and republish.</td>
</tr>""")

# ---------------------------------------------------------------------------
# 10. Write 5-Tab Handoff HTML Report (Automated, Manual, 301 Redirects, SEO & AIO Strategy, Wave/Media/L10n)
# ---------------------------------------------------------------------------
w1_s = wave_summary["w1"]
w2_s = wave_summary["w2"]
w3_s = wave_summary["w3"]

full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ServiceNow Rebrand Handoff Dashboard: Super Intelligence (SI) Control Tower</title>
<style>
  :root {{
    --bg: #08121a;
    --surface: #0f1e29;
    --surface-2: #162936;
    --border: #233b4d;
    --green: #62d84e;
    --green-dim: rgba(98, 216, 78, 0.14);
    --text: #f0f6fa;
    --muted: #9ab0c2;
    --amber: #f59e0b;
    --amber-dim: rgba(245, 158, 11, 0.15);
    --red: #ef4444;
    --red-dim: rgba(239, 68, 68, 0.15);
    --blue: #38bdf8;
    --blue-dim: rgba(56, 189, 248, 0.15);
    --purple: #c084fc;
    --purple-dim: rgba(192, 132, 252, 0.15);
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    max-width: 100%;
    overflow-x: hidden;
  }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.55;
    padding-bottom: 60px;
  }}
  a, .ext-link {{
    color: var(--blue);
    text-decoration: underline;
    text-underline-offset: 2px;
    word-break: break-word;
    overflow-wrap: anywhere;
    cursor: pointer;
  }}
  a:hover, .ext-link:hover {{
    color: var(--green);
  }}
  header {{
    background: linear-gradient(135deg, #09151e 0%, #102a3b 100%);
    border-bottom: 1px solid var(--border);
    padding: clamp(14px, 2.2vw, 24px) clamp(14px, 3vw, 40px);
    position: sticky;
    top: 0;
    z-index: 100;
    backdrop-filter: blur(12px);
  }}
  .header-top {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
    max-width: 1440px;
    margin: 0 auto;
  }}
  .brand-pill {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: var(--green-dim);
    color: var(--green);
    border: 1px solid rgba(98, 216, 78, 0.4);
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 11.5px;
    font-weight: 700;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    margin-bottom: 6px;
    max-width: 100%;
    white-space: normal;
  }}
  h1 {{ font-size: clamp(18px, 2.2vw, 24px); font-weight: 700; letter-spacing: -0.02em; line-height: 1.25; }}
  h1 span {{ color: var(--green); }}
  .subtitle {{ color: var(--muted); font-size: clamp(12px, 1.4vw, 14px); margin-top: 4px; }}
  .container {{ max-width: 1440px; margin: 20px auto; padding: 0 clamp(12px, 3vw, 40px); }}
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(185px, 1fr));
    gap: 12px;
    margin-bottom: 20px;
  }}
  .kpi-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 14px 16px;
    min-width: 0;
  }}
  .kpi-label {{ color: var(--muted); font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; font-weight: 600; }}
  .kpi-value {{ font-size: clamp(19px, 2vw, 24px); font-weight: 800; margin: 5px 0 4px; color: var(--text); }}
  .kpi-value.green {{ color: var(--green); }}
  .kpi-value.amber {{ color: var(--amber); }}
  .kpi-value.blue {{ color: var(--blue); }}
  .kpi-value.purple {{ color: var(--purple); }}
  .kpi-sub {{ font-size: 12px; color: var(--muted); overflow-wrap: anywhere; }}
  .tabs {{
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 20px;
    padding-bottom: 6px;
  }}
  .tab-btn {{
    background: var(--surface);
    color: var(--muted);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px 14px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s;
    text-align: left;
    line-height: 1.35;
    flex: 1 1 210px;
  }}
  .tab-btn:hover {{ color: var(--text); background: rgba(255,255,255,0.04); }}
  .tab-btn.active {{ color: var(--green); border-color: var(--green); background: var(--green-dim); }}
  .tab-panel {{ display: none; }}
  .tab-panel.active {{ display: block; }}
  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: clamp(14px, 2vw, 22px);
    margin-bottom: 20px;
    min-width: 0;
  }}
  .card h2 {{ font-size: clamp(15px, 1.8vw, 18px); margin-bottom: 6px; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; line-height: 1.35; }}
  .card p.desc {{ color: var(--muted); font-size: 13px; margin-bottom: 14px; }}
  .alert {{
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 16px;
    border-left: 4px solid var(--green);
    background: var(--surface-2);
    font-size: 13px;
    overflow-wrap: anywhere;
  }}
  .alert.amber {{ border-left-color: var(--amber); background: var(--amber-dim); }}
  .alert.blue {{ border-left-color: var(--blue); background: var(--blue-dim); }}
  .alert strong {{ display: block; margin-bottom: 4px; font-size: 14px; }}
  .controls {{
    display: flex;
    gap: 6px;
    margin-bottom: 10px;
    flex-wrap: wrap;
    align-items: center;
  }}
  .search-input {{
    flex: 1 1 240px;
    min-width: 0;
    width: 100%;
    background: var(--bg);
    border: 1px solid var(--border);
    color: var(--text);
    padding: 9px 12px;
    border-radius: 8px;
    font-size: 13px;
  }}
  .filter-btn {{
    background: var(--surface-2);
    border: 1px solid var(--border);
    color: var(--muted);
    padding: 6px 10px;
    border-radius: 8px;
    font-size: 11.5px;
    font-weight: 600;
    cursor: pointer;
    line-height: 1.3;
  }}
  .filter-btn.active, .filter-btn:hover {{
    background: var(--green-dim);
    color: var(--green);
    border-color: var(--green);
  }}
  .table-wrap {{
    width: 100%;
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    table-layout: auto;
  }}
  th, td {{
    text-align: left;
    padding: 11px 12px;
    border-bottom: 1px solid var(--border);
    vertical-align: top;
    overflow-wrap: anywhere;
    word-break: break-word;
  }}
  th {{
    background: var(--surface-2);
    color: var(--muted);
    font-weight: 600;
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: 0.05em;
  }}
  tr:hover td {{ background: rgba(255,255,255,0.02); }}
  .badge {{
    display: inline-block;
    padding: 3px 7px;
    border-radius: 6px;
    font-size: 11px;
    font-weight: 700;
    white-space: normal;
    line-height: 1.3;
  }}
  .b-w1 {{ background: var(--red-dim); color: #fca5a5; border: 1px solid rgba(239,68,68,0.3); }}
  .b-w2 {{ background: var(--amber-dim); color: #fcd34d; border: 1px solid rgba(245,158,11,0.3); }}
  .b-w3 {{ background: var(--blue-dim); color: #7dd3fc; border: 1px solid rgba(56,189,248,0.3); }}
  .b-ok {{ background: var(--green-dim); color: var(--green); border: 1px solid rgba(98,216,78,0.3); }}
  .b-warn {{ background: var(--purple-dim); color: #d8b4fe; border: 1px solid rgba(192,132,252,0.3); }}
  code, .mono {{
    font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
    font-size: 12px;
    background: rgba(0,0,0,0.35);
    padding: 2px 5px;
    border-radius: 4px;
    color: #93c5fd;
    overflow-wrap: anywhere;
    word-break: break-word;
    white-space: normal;
  }}
  pre {{
    background: #060d13;
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 12px;
    overflow-x: auto;
    white-space: pre-wrap;
    word-break: break-word;
    font-family: "SFMono-Regular", Consolas, monospace;
    font-size: 11.5px;
    color: #e2e8f0;
    margin: 10px 0;
  }}

  @media (max-width: 1120px) {{
    header {{
      position: relative;
    }}
    .kpi-grid {{
      grid-template-columns: repeat(auto-fit, minmax(155px, 1fr));
    }}
    table, thead, tbody, th, td, tr {{
      display: block;
      width: 100%;
    }}
    thead tr {{
      position: absolute;
      top: -9999px;
      left: -9999px;
    }}
    tbody tr {{
      background: var(--surface-2);
      border: 1px solid var(--border);
      border-radius: 10px;
      margin-bottom: 14px;
      padding: 10px 12px;
    }}
    tbody td {{
      border-bottom: 1px solid rgba(255,255,255,0.06);
      padding: 8px 4px;
      width: 100% !important;
    }}
    tbody td:last-child {{
      border-bottom: none;
    }}
    tbody td[data-label]::before {{
      content: attr(data-label);
      display: block;
      font-size: 10.5px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: var(--muted);
      margin-bottom: 4px;
    }}
  }}

  @media (max-width: 640px) {{
    .kpi-grid {{
      grid-template-columns: 1fr 1fr;
    }}
    .tab-btn {{
      flex: 1 1 100%;
    }}
    .filter-btn {{
      flex: 1 1 auto;
      text-align: center;
    }}
  }}
</style>
<script>
// Tab Switching
window.showTab = function(tabId, btn) {{
  document.querySelectorAll('.tab-panel').forEach(function(p) {{ p.classList.remove('active'); }});
  document.querySelectorAll('.tab-btn').forEach(function(b) {{ b.classList.remove('active'); }});
  var panel = document.getElementById(tabId);
  if (panel) panel.classList.add('active');
  if (btn) btn.classList.add('active');
}};

// Tab 1: Automated Updates Filters
window.autoCat = 'all';
window.autoExec = 'all';
window.autoLoc = 'all';

window.filterAutoCat = function(cat, btn) {{
  window.autoCat = cat;
  if (btn && btn.parentElement) {{
    btn.parentElement.querySelectorAll('.filter-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
  }}
  window.filterAutoTable();
}};

window.filterAutoLoc = function(loc, btn) {{
  window.autoLoc = loc;
  if (btn && btn.parentElement) {{
    btn.parentElement.querySelectorAll('.filter-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
  }}
  window.filterAutoTable();
}};

window.filterAutoExec = function(execMode, btn) {{
  window.autoExec = execMode;
  if (btn && btn.parentElement) {{
    btn.parentElement.querySelectorAll('.filter-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
  }}
  window.filterAutoTable();
}};

window.filterAutoTable = function() {{
  var searchEl = document.getElementById('autoSearch');
  var q = searchEl ? searchEl.value.toLowerCase() : '';
  var rows = document.querySelectorAll('#autoTable tbody tr');
  rows.forEach(function(r) {{
    var text = r.innerText.toLowerCase();
    var cat = r.getAttribute('data-cat');
    var b30 = r.getAttribute('data-b30');
    var ex = r.getAttribute('data-exec');
    var loc = r.getAttribute('data-loc');
    var matchQ = !q || text.includes(q);
    var matchCat = true;
    if (window.autoCat === 'w1' || window.autoCat === 'w2' || window.autoCat === 'w3') {{
      matchCat = (cat === window.autoCat);
    }} else if (window.autoCat && window.autoCat.startsWith('cat-')) {{
      matchCat = (b30 === window.autoCat);
    }}
    var matchEx = (window.autoExec === 'all') || (ex === window.autoExec);
    var matchLoc = true;
    if (window.autoLoc === 'us') {{
      matchLoc = (loc === 'us');
    }} else if (window.autoLoc === 'regional') {{
      matchLoc = (loc !== 'us');
    }} else if (window.autoLoc !== 'all') {{
      matchLoc = (loc === window.autoLoc);
    }}
    r.style.display = (matchQ && matchCat && matchEx && matchLoc) ? '' : 'none';
  }});
}};

// Tab 2: Manual Updates Filters
window.manCat = 'all';
window.manType = 'all';
window.manLoc = 'all';

window.filterManCat = function(cat, btn) {{
  window.manCat = cat;
  if (btn && btn.parentElement) {{
    btn.parentElement.querySelectorAll('.filter-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
  }}
  window.filterManTable();
}};

window.filterManLoc = function(loc, btn) {{
  window.manLoc = loc;
  if (btn && btn.parentElement) {{
    btn.parentElement.querySelectorAll('.filter-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
  }}
  window.filterManTable();
}};

window.filterManType = function(mtype, btn) {{
  window.manType = mtype;
  if (btn && btn.parentElement) {{
    btn.parentElement.querySelectorAll('.filter-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
  }}
  window.filterManTable();
}};

window.filterManTable = function() {{
  var searchEl = document.getElementById('manSearch');
  var q = searchEl ? searchEl.value.toLowerCase() : '';
  var rows = document.querySelectorAll('#manTable tbody tr');
  rows.forEach(function(r) {{
    var text = r.innerText.toLowerCase();
    var cat = r.getAttribute('data-cat');
    var mtypes = r.getAttribute('data-mtypes') || '';
    var loc = r.getAttribute('data-loc');
    var matchQ = !q || text.includes(q);
    var matchCat = (window.manCat === 'all') || (cat === window.manCat);
    var matchType = (window.manType === 'all') || mtypes.includes(window.manType);
    var matchLoc = true;
    if (window.manLoc === 'us') {{
      matchLoc = (loc === 'us');
    }} else if (window.manLoc === 'regional') {{
      matchLoc = (loc !== 'us');
    }} else if (window.manLoc !== 'all') {{
      matchLoc = (loc === window.manLoc);
    }}
    r.style.display = (matchQ && matchCat && matchType && matchLoc) ? '' : 'none';
  }});
}};

// Tab 3: 301 Redirects Filters
window.redType = 'all';
window.filterRedType = function(rtype, btn) {{
  window.redType = rtype;
  if (btn && btn.parentElement) {{
    btn.parentElement.querySelectorAll('.filter-btn').forEach(function(b) {{ b.classList.remove('active'); }});
    btn.classList.add('active');
  }}
  window.filterRedTable();
}};

window.filterRedTable = function() {{
  var searchEl = document.getElementById('redSearch');
  var q = searchEl ? searchEl.value.toLowerCase() : '';
  var rows = document.querySelectorAll('#redTable tbody tr');
  rows.forEach(function(r) {{
    var text = r.innerText.toLowerCase();
    var rtype = r.getAttribute('data-rtype');
    var matchQ = !q || text.includes(q);
    var matchType = (window.redType === 'all') || (rtype === window.redType);
    r.style.display = (matchQ && matchType) ? '' : 'none';
  }});
}};
</script>
</head>
<body>

<header>
  <div class="header-top">
    <div>
      <div class="brand-pill">ServiceNow Digital UX • Global + Regional Rebrand Execution Handoff</div>
      <h1>AI Control Tower &rarr; <span>Super Intelligence (SI) Control Tower</span></h1>
      <div class="subtitle">{len(unique_body_pages)} Total Body-Match Pages ({us_pages_total} US + {regional_pages_total} Regional Across 13 Locales, incl. {a_link('https://www.servicenow.com/uk/', '/uk/')}) • Separate Automated vs. Manual Queues • 301 Redirects • SEO &amp; AIO Strategy</div>
    </div>
  </div>
</header>

<div class="container">
  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-label">Total Rebrand Scope</div>
      <div class="kpi-value green">{len(unique_body_pages)} Pages</div>
      <div class="kpi-sub"><strong>🇺🇸 {us_pages_total} US Pages</strong> + <strong>🌍 {regional_pages_total} Regional Pages</strong> across 13 locales (0 pages dropped)</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Wave 1 Impact ({w1_s['pages']} Pages)</div>
      <div class="kpi-value green">96.2% Traffic</div>
      <div class="kpi-sub"><strong>{w1_s['us_pages']} US + {w1_s['reg_pages']} Regional</strong> (incl. {a_link('https://www.servicenow.com/uk/', '/uk/')}, {a_link('https://www.servicenow.com/de/', '/de/')}, {a_link('https://www.servicenow.com/fr/', '/fr/')}) • <strong>{w1_s['visits']:,} visits</strong></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">🤖 Automated Updates (Tab 1)</div>
      <div class="kpi-value blue">{total_auto_instances:,} Swaps</div>
      <div class="kpi-sub">Across all <strong>{len(unique_body_pages)} pages</strong> (<strong>{pages_100pct_auto} pure auto</strong> + <strong>{pages_needing_manual} hybrid</strong>) + {len(ssi_only_pages):,} Global Nav XF pages</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">🛠️ Manual Updates (Tab 2)</div>
      <div class="kpi-value amber">{pages_needing_manual} Pages</div>
      <div class="kpi-sub"><strong>{total_manual_copy_instances} WWSG copy edits</strong>, 34 page slug moves, <strong>{total_baked_images_count} baked-in images ({total_baked_designs_count} visual designs)</strong>, 13 PDFs &amp; 10 L10n reviews</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">🔀 301 Redirects (Tab 3)</div>
      <div class="kpi-value purple">{total_redirects_count} Redirects</div>
      <div class="kpi-sub"><strong>{webpage_redirect_cnt} Webpage slugs</strong> + <strong>{pdf_redirect_cnt} DAM PDFs</strong> + <strong>{img_redirect_cnt} DAM Images</strong> via HTTP 301</div>
    </div>
  </div>

  <div class="tabs">
    <button id="btn-tab-automated" class="tab-btn active" onclick="showTab('tab-automated', this)">1. 🤖 Automated Updates ({len(unique_body_pages)} Pages • {total_auto_instances:,} Bulk Swaps)</button>
    <button id="btn-tab-manual" class="tab-btn" onclick="showTab('tab-manual', this)">2. 🛠️ Manual Updates ({pages_needing_manual} Pages • {total_manual_tasks_all} Manual Tasks)</button>
    <button id="btn-tab-redirects" class="tab-btn" onclick="showTab('tab-redirects', this)">3. 🔀 301 Redirects ({total_redirects_count} URL &amp; DAM Slugs)</button>
    <button id="btn-tab-seo" class="tab-btn" onclick="showTab('tab-seo', this)">4. 🔍 SEO &amp; AIO Strategy (SI vs. Super Intelligence)</button>
    <button id="btn-tab-waves" class="tab-btn" onclick="showTab('tab-waves', this)">5. 📊 Wave Rollout &amp; Media/L10n Summary</button>
  </div>

  <!-- TAB 1: AUTOMATED UPDATES -->
  <div id="tab-automated" class="tab-panel active">
    <div class="alert blue">
      <strong>🤖 Automated Bulk Update Queue (All {len(unique_body_pages)} Pages Have Automated Instances — Including the {pages_needing_manual} Hybrid Pages!):</strong>
      This tab lists ONLY the <strong>100% automated DOM, metadata, link, alt-text, and 0-pixel DAM image swaps</strong> executed by the AEM Groovy script across all {len(unique_body_pages)} pages ({us_pages_total} US + {regional_pages_total} Regional). For the <strong>{pages_needing_manual} hybrid pages</strong> (such as {a_link('https://www.servicenow.com/products/ai-control-tower.html')}, {a_link('https://www.servicenow.com/platform/responsible-ai.html')}, and {a_link('https://www.servicenow.com/fr/products/ai-control-tower.html')}) that have <em>both</em> automated bulk swaps and manual tasks, their automated instances are shown here in Tab 1 (including the H1 Hero <code>Super Intelligence Control Tower</code> swap), while their manual tasks are itemized separately in <strong>Tab 2 (🛠️ Manual Updates)</strong>!
    </div>

    <div class="card">
      <h2>Tab 1: Automated Bulk Updates ({len(unique_body_pages)} Pages: {pages_100pct_auto} Pure Automated + {pages_needing_manual} Hybrid Pages)</h2>
      <p class="desc">Filter by <strong>Wave</strong>, <strong>Region / Locale</strong>, <strong>Automation Scope</strong> (Pure Auto vs. Hybrid URLs that also appear in Tab 2), or <strong>Page Category</strong>.</p>

      <div class="controls">
        <input type="text" id="autoSearch" class="search-input" placeholder="Search Automated Queue by URL (e.g. /uk/, /de/, /fr/), title, automated copy snippet, 'SI Control Tower'..." onkeyup="filterAutoTable()">
        <button class="filter-btn active" onclick="filterAutoCat('all', this)">All Pages ({len(unique_body_pages)})</button>
        <button class="filter-btn" onclick="filterAutoCat('w1', this)">Wave 1 ({w1_s['pages']} pages • 96.2% Traffic)</button>
        <button class="filter-btn" onclick="filterAutoCat('w2', this)">Wave 2 ({w2_s['pages']} pages • 3.6% Traffic)</button>
        <button class="filter-btn" onclick="filterAutoCat('w3', this)">Wave 3 ({w3_s['pages']} pages • 0.2% Traffic)</button>
      </div>
      <div class="controls">
        <span style="font-size:12px;color:var(--muted);font-weight:600;margin-right:4px;">Region / Locale:</span>
        <button class="filter-btn active" onclick="filterAutoLoc('all', this)">All Regions ({len(unique_body_pages)})</button>
        <button class="filter-btn" onclick="filterAutoLoc('us', this)">🇺🇸 US ({locale_counts['us']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('regional', this)">🌍 All Regional ({regional_pages_total})</button>
        <button class="filter-btn" onclick="filterAutoLoc('uk', this)">🇬🇧 UK ({locale_counts['uk']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('au', this)">🇦🇺 AU ({locale_counts['au']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('in', this)">🇮🇳 IN ({locale_counts['in']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('de', this)">🇩🇪 DE ({locale_counts['de']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('fr', this)">🇫🇷 FR ({locale_counts['fr']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('fr-ca', this)">🇨🇦 FR-CA ({locale_counts['fr-ca']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('es', this)">🇪🇸 ES ({locale_counts['es']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('latam', this)">🌎 LATAM ({locale_counts['latam']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('br', this)">🇧🇷 BR ({locale_counts['br']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('it', this)">🇮🇹 IT ({locale_counts['it']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('nl', this)">🇳🇱 NL ({locale_counts['nl']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('jp', this)">🇯🇵 JP ({locale_counts['jp']})</button>
        <button class="filter-btn" onclick="filterAutoLoc('kr', this)">🇰🇷 KR ({locale_counts['kr']})</button>
      </div>
      <div class="controls">
        <span style="font-size:12px;color:var(--muted);font-weight:600;margin-right:4px;">Automation Scope:</span>
        <button class="filter-btn active" onclick="filterAutoExec('all', this)">All Automated Pages ({len(unique_body_pages)})</button>
        <button class="filter-btn" onclick="filterAutoExec('pure-auto', this)">🤖 100% Pure Automated ({pages_100pct_auto})</button>
        <button class="filter-btn" onclick="filterAutoExec('hybrid-auto', this)">🔀 Hybrid URLs (Also Have Manual Tasks in Tab 2: {pages_needing_manual})</button>
        <span style="font-size:12px;color:var(--muted);font-weight:600;margin-left:6px;margin-right:4px;">Category:</span>
        <button class="filter-btn" onclick="filterAutoCat('cat-product', this)">Product</button>
        <button class="filter-btn" onclick="filterAutoCat('cat-solutions', this)">Solutions &amp; Apps</button>
        <button class="filter-btn" onclick="filterAutoCat('cat-platform', this)">Homepage &amp; Platform</button>
        <button class="filter-btn" onclick="filterAutoCat('cat-events', this)">Events</button>
        <button class="filter-btn" onclick="filterAutoCat('cat-resources', this)">Resources &amp; ePDFs</button>
        <button class="filter-btn" onclick="filterAutoCat('cat-editorial', this)">Workflow &amp; Blogs</button>
      </div>

      <div class="table-wrap">
        <table id="autoTable">
          <thead>
            <tr>
              <th style="width:14%;">Wave, Region &amp; Track</th>
              <th style="width:20%;">Full Clickable Page URL</th>
              <th style="width:15%;">Traffic &amp; Inclusion Basis</th>
              <th style="width:38%;">Automated Instances Only (🤖 Groovy Bulk Swap: Meta, Body, Links, Alt &amp; 0-Pixel DAM)</th>
              <th style="width:13%;">Automated Summary</th>
            </tr>
          </thead>
          <tbody>
            {"".join(auto_table_rows_html)}
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- TAB 2: MANUAL UPDATES -->
  <div id="tab-manual" class="tab-panel">
    <div class="alert amber">
      <strong>🛠️ Manual Update Queue ({pages_needing_manual} Pages with {total_manual_tasks_all} Total Manual Tasks — Separated from Automated Swaps!):</strong>
      Every page listed below ALSO has automated bulk swaps in <strong>Tab 1 (🤖 Automated Updates)</strong> that run first via the AEM Groovy script. This tab isolates ONLY the <strong>manual tasks</strong> required on these {pages_needing_manual} pages: <strong>{total_manual_copy_instances} WWSG Copy Edits</strong> (where pre-existing flagship FAQ/Marquee copy exceeded 50/175/250 chars), <strong>34 AEM Page Node Slug Moves &amp; 301s</strong>, <strong>{total_baked_images_count} Baked-In Image Re-renders ({total_baked_designs_count} distinct visual designs</strong> verified via Apple Vision OCR across 388 pages + 8,113 Brand 3.0 DAM assets, including <code>welcome-ai-control-tower-.sm.png</code> on <code>/platform/responsible-ai.html</code>), <strong>13 Live/Vector PDF Updates</strong>, and <strong>10 Non-English Flagship L10n Grammar Reviews</strong>.
    </div>

    <div class="card">
      <h2>Tab 2: Manual Updates Queue ({pages_needing_manual} Pages • {total_manual_tasks_all} Manual Tasks)</h2>
      <p class="desc">Filter by <strong>Manual Task Type</strong> (<code>Copy Edit</code>, <code>301 Slug</code>, <code>Baked-In Image</code>, <code>PDF Update</code>, <code>L10n Review</code>), <strong>Wave</strong>, or <strong>Region / Locale</strong>.</p>

      <div class="controls">
        <input type="text" id="manSearch" class="search-input" placeholder="Search Manual Queue by URL, WWSG rewrite, PDF filename, image filename, locale..." onkeyup="filterManTable()">
        <button class="filter-btn active" onclick="filterManCat('all', this)">All Manual Pages ({pages_needing_manual})</button>
        <button class="filter-btn" onclick="filterManCat('w1', this)">Wave 1 ({manual_wave_counts['w1']} pages)</button>
        <button class="filter-btn" onclick="filterManCat('w2', this)">Wave 2 ({manual_wave_counts['w2']} pages)</button>
        <button class="filter-btn" onclick="filterManCat('w3', this)">Wave 3 ({manual_wave_counts['w3']} pages)</button>
      </div>
      <div class="controls">
        <span style="font-size:12px;color:var(--muted);font-weight:600;margin-right:4px;">Manual Task Type:</span>
        <button class="filter-btn active" onclick="filterManType('all', this)">All Task Types ({pages_needing_manual} pages)</button>
        <button class="filter-btn" onclick="filterManType('m-copy', this)">✍️ Copy Edit: WWSG Trim ({manual_type_counts['m-copy']} pages / {total_manual_copy_instances} instances)</button>
        <button class="filter-btn" onclick="filterManType('m-slug', this)">🔀 301 URL Slug Move ({manual_type_counts['m-slug']} pages)</button>
        <button class="filter-btn" onclick="filterManType('m-img', this)">🖼️ Baked-In Image Re-render ({manual_type_counts['m-img']} pages)</button>
        <button class="filter-btn" onclick="filterManType('m-pdf', this)">📄 Live/Vector PDF Update ({manual_type_counts['m-pdf']} pages)</button>
        <button class="filter-btn" onclick="filterManType('m-l10n', this)">🌐 L10n Grammar Review ({manual_type_counts['m-l10n']} pages)</button>
      </div>
      <div class="controls">
        <span style="font-size:12px;color:var(--muted);font-weight:600;margin-right:4px;">Region / Locale:</span>
        <button class="filter-btn active" onclick="filterManLoc('all', this)">All Regions ({pages_needing_manual})</button>
        <button class="filter-btn" onclick="filterManLoc('us', this)">🇺🇸 US ({manual_locale_counts['us']})</button>
        <button class="filter-btn" onclick="filterManLoc('regional', this)">🌍 All Regional ({pages_needing_manual - manual_locale_counts['us']})</button>
        <button class="filter-btn" onclick="filterManLoc('uk', this)">🇬🇧 UK ({manual_locale_counts['uk']})</button>
        <button class="filter-btn" onclick="filterManLoc('au', this)">🇦🇺 AU ({manual_locale_counts['au']})</button>
        <button class="filter-btn" onclick="filterManLoc('in', this)">🇮🇳 IN ({manual_locale_counts['in']})</button>
        <button class="filter-btn" onclick="filterManLoc('de', this)">🇩🇪 DE ({manual_locale_counts['de']})</button>
        <button class="filter-btn" onclick="filterManLoc('fr', this)">🇫🇷 FR ({manual_locale_counts['fr']})</button>
        <button class="filter-btn" onclick="filterManLoc('fr-ca', this)">🇨🇦 FR-CA ({manual_locale_counts['fr-ca']})</button>
        <button class="filter-btn" onclick="filterManLoc('es', this)">🇪🇸 ES ({manual_locale_counts['es']})</button>
        <button class="filter-btn" onclick="filterManLoc('latam', this)">🌎 LATAM ({manual_locale_counts['latam']})</button>
        <button class="filter-btn" onclick="filterManLoc('br', this)">🇧🇷 BR ({manual_locale_counts['br']})</button>
        <button class="filter-btn" onclick="filterManLoc('it', this)">🇮🇹 IT ({manual_locale_counts['it']})</button>
        <button class="filter-btn" onclick="filterManLoc('nl', this)">🇳🇱 NL ({manual_locale_counts['nl']})</button>
        <button class="filter-btn" onclick="filterManLoc('jp', this)">🇯🇵 JP ({manual_locale_counts['jp']})</button>
        <button class="filter-btn" onclick="filterManLoc('kr', this)">🇰🇷 KR ({manual_locale_counts['kr']})</button>
      </div>

      <div class="table-wrap">
        <table id="manTable">
          <thead>
            <tr>
              <th style="width:14%;">Wave, Region &amp; Track</th>
              <th style="width:20%;">Full Clickable Page URL</th>
              <th style="width:15%;">Traffic &amp; Inclusion Basis</th>
              <th style="width:38%;">Manual Tasks Only (🛠️ WWSG Copy Edits, 301 Slug, Baked-In Images, PDFs &amp; L10n)</th>
              <th style="width:13%;">Manual Summary</th>
            </tr>
          </thead>
          <tbody>
            {"".join(manual_table_rows_html)}
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- TAB 3: 301 REDIRECTS -->
  <div id="tab-redirects" class="tab-panel">
    <div class="alert">
      <strong>🔀 Complete HTTP 301 Permanent Redirect Inventory ({total_redirects_count} Total URL &amp; DAM Slug Migrations):</strong>
      To preserve 100% of organic search equity, external backlinks, bookmark links, and embedded asset paths during the clean cutover to <strong>Super Intelligence Control Tower</strong>, deploy HTTP 301 permanent redirects across all <strong>{webpage_redirect_cnt} Webpage URL slugs</strong> (US + 13 regional locales), <strong>{pdf_redirect_cnt} DAM PDF filenames</strong>, and <strong>{img_redirect_cnt} DAM Image filenames</strong>.
    </div>

    <div class="card">
      <h2>Part 1: Itemized 301 Redirect Mapping Table ({total_redirects_count} Redirects: {webpage_redirect_cnt} Webpages + {pdf_redirect_cnt} PDFs + {img_redirect_cnt} Images)</h2>
      <div class="controls">
        <input type="text" id="redSearch" class="search-input" placeholder="Search 301 Redirects by old URL, new URL, locale (/uk/, /de/, /fr/), .pdf, .png..." onkeyup="filterRedTable()">
        <button class="filter-btn active" onclick="filterRedType('all', this)">All 301 Redirects ({total_redirects_count})</button>
        <button class="filter-btn" onclick="filterRedType('webpage', this)">🌐 Webpage Slugs ({webpage_redirect_cnt})</button>
        <button class="filter-btn" onclick="filterRedType('pdf', this)">📄 DAM PDF Slugs ({pdf_redirect_cnt})</button>
        <button class="filter-btn" onclick="filterRedType('image', this)">🖼️ DAM Image Slugs ({img_redirect_cnt})</button>
      </div>
      <div class="table-wrap">
        <table id="redTable">
          <thead>
            <tr>
              <th style="width:6%;">#</th>
              <th style="width:16%;">Asset Type, Wave &amp; Locale</th>
              <th style="width:32%;">Old Live URL (Clickable 301 Source)</th>
              <th style="width:30%;">New Target URL (301 Destination)</th>
              <th style="width:16%;">Status &amp; Technical Action</th>
            </tr>
          </thead>
          <tbody>
            {"".join(redirect_rows_html)}
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Part 2: Ready-to-Deploy Apache Dispatcher (<code>mod_rewrite</code>) &amp; Akamai Edge Redirector Rules</h2>
      <p class="desc">Drop-in configuration block handling all US + 13 regional webpage slugs, DAM PDFs (including <code>.coredownload</code>), and DAM images with a single regex ruleset.</p>
      <pre># ==============================================================================
# ServiceNow Apache Dispatcher / Akamai Edge Redirector — 301 Cutover Rules
# Product Rename: AI Control Tower (AICT) -> Super Intelligence Control Tower
# ==============================================================================
RewriteEngine On

# 1. Flagship Product Page, Pricing Subpage + All 13 Regional Locales (/uk, /au, /in, /de, /fr, /fr-ca, /es, /latam, /br, /it, /nl, /jp, /kr)
RewriteRule ^(/([a-z]{{2}}(-[a-z]{{2}})?|latam))?/products/ai-control-tower(/pricing)?\\.html$ $1/products/super-intelligence-control-tower$4.html [R=301,L,NC]

# 2. Contact Us, Webinar Promo & Discovery Playbook eBook (aict / ai-control-tower slugs)
RewriteRule ^/contact-us-aict\\.html$ /contact-us-super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/lpwbr/meet-ai-control-tower-enterprise-ai-governance-at-machine-speed\\.html$ /lpwbr/meet-super-intelligence-control-tower-enterprise-ai-governance-at-machine-speed.html [R=301,L,NC]
RewriteRule ^/standard/resource-center/ebook/ebk-aict-discovery-playbook\\.html$ /standard/resource-center/ebook/ebk-super-intelligence-control-tower-discovery-playbook.html [R=301,L,NC]

# 3. Live Events, Knowledge Sessions & On-Demand Webinars (US + 13 Regional Locales)
RewriteRule ^(/([a-z]{{2}}(-[a-z]{{2}})?|latam))?/events/live/ai-control-tower\\.html$ $1/events/live/super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/events/on-demand-webinars/put-ai-to-work-with-ai-control-tower\\.html$ /events/on-demand-webinars/put-ai-to-work-with-super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/events/knowledge/2024/sessions/now-on-now--turning-ai-governance-risk-into-reward-with-the-ai-control-tower\\.html$ /events/knowledge/2024/sessions/now-on-now--turning-ai-governance-risk-into-reward-with-the-super-intelligence-control-tower.html [R=301,L,NC]

# 4. Resource Center White Papers, Solution Briefs & eBooks (US + DE/FR/UK)
RewriteRule ^(/([a-z]{{2}}(-[a-z]{{2}})?|latam))?/standard/resource-center/(white-paper|solution-brief|ebook)/(.*)ai-control-tower(.*)\\.html$ $1/standard/resource-center/$4/$5super-intelligence-control-tower$6.html [R=301,L,NC]

# 5. Workflow Articles, Regional Blogs & Documentation Landing Page
RewriteRule ^/workflow/ai/(.*)ai-control-tower\\.html$ /workflow/ai/$1super-intelligence-control-tower.html [R=301,L,NC]
RewriteRule ^/(uk|de|nl)/blogs/2025/(.*)ai-control-tower$ /$1/blogs/2025/$2super-intelligence-control-tower [R=301,L,NC]
RewriteRule ^/docs/r/intelligent-experiences/ai-control-tower/ai-control-tower-landing\\.html$ /docs/r/intelligent-experiences/super-intelligence-control-tower/super-intelligence-control-tower-landing.html [R=301,L,NC]

# 6. DAM PDFs & .coredownload Endpoints (EN-US, DE-DE, FR-FR)
RewriteRule ^/content/dam/servicenow-assets/public/(.*)/ebk-aict-discovery-playbook\\.pdf(\\.coredownload)?$ /content/dam/servicenow-assets/public/$1/ebk-super-intelligence-control-tower-discovery-playbook.pdf$2 [R=301,L,NC]
RewriteRule ^/content/dam/servicenow-assets/public/(.*)/(.*)ai-control-tower(.*)\\.pdf(\\.coredownload)?$ /content/dam/servicenow-assets/public/$1/$2super-intelligence-control-tower$3.pdf$4 [R=301,L,NC]

# 7. DAM Images (.png, .jpg, .webp — excludes partner logo-aicte.svg)
RewriteCond %{{REQUEST_URI}} !logo-aicte\\.svg$ [NC]
RewriteRule ^/content/dam/servicenow-assets/public/(.*)/(.*)ai-control-tower(.*)\\.(png|jpg|jpeg|webp)$ /content/dam/servicenow-assets/public/$1/$2super-intelligence-control-tower$3.$4 [R=301,L,NC]</pre>
    </div>
  </div>

  <!-- TAB 4: SEO & AIO STRATEGY (SI VS. SUPER INTELLIGENCE) -->
  <div id="tab-seo" class="tab-panel">
    <div class="alert blue">
      <strong>🔍 Semrush MCP + Live SERP &amp; AIO (AI Overview / LLM Citation) Optimization Audit:</strong>
      Connected to the <strong>Semrush MCP Server</strong> (<code>https://mcp.semrush.com/v2/mcp</code> via OAuth 2.0 PKCE authentication; corporate sub-account API unit status verified alongside full live DOM, Google SERP pixel-width, Schema.org, and LLM retrieval-augmented generation [RAG] entity analysis across all {len(unique_body_pages)} pages). Below is the complete SEO &amp; AIO decision framework governing when to use <strong><code>Super Intelligence Control Tower</code> (32 chars)</strong> vs. <strong><code>SI Control Tower</code> (16 chars)</strong>.
    </div>

    <div class="card">
      <h2>Part 1: Entity Canonicalization for AIO (ChatGPT / Perplexity / Google AI Overviews) vs. SERP Pixel Limits</h2>
      <p class="desc">Why pairing the 32-char Full Canonical Name on primary anchors with the 16-char <code>SI Control Tower</code> shorthand on constrained UI components maximizes both LLM entity citation and Google SERP click-through rate.</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th style="width:24%;">Optimization Pillar</th>
              <th style="width:38%;">Technical Mechanism (Google Search &amp; LLM / AIO Retrieval)</th>
              <th style="width:38%;">Exact Implementation Rule Across the {len(unique_body_pages)} Pages</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td data-label="Optimization Pillar"><span class="badge b-ok">1. AIO &amp; LLM Entity Resolution</span><br><strong>Primary Canonical Anchors Use Full Name (32 Chars)</strong></td>
              <td data-label="Technical Mechanism">
                • Generative search engines (<strong>Google AI Overviews, ChatGPT Search, Perplexity, Gemini, Claude</strong>) and Google's Knowledge Graph extract the canonical product entity from 5 primary DOM &amp; structured-data signals: <code>&lt;title&gt;</code>, <code>&lt;h1&gt;</code>, First Body Paragraph, Signature Company Tagline, and JSON-LD <code>SoftwareApplication.name</code>.<br>
                • Using the full name <strong><code>Super Intelligence Control Tower</code></strong> in these 5 primary anchors ensures LLMs index the full semantic phrase <em>"Super Intelligence"</em> alongside <em>"Control Tower"</em>.
              </td>
              <td data-label="Exact Implementation Rule">
                • <strong>Flagship <code>&lt;title&gt;</code>:</strong> <code>Super Intelligence Control Tower - ServiceNow</code> (45/60 chars, ~412px / 600px SERP cap ✓).<br>
                • <strong>Flagship <code>&lt;h1&gt;</code>:</strong> <code>Super Intelligence Control Tower</code> (32/50 chars ✓ — 100% Automated).<br>
                • <strong>First Body Mention &amp; Meta Description:</strong> Introduce as <code>Super Intelligence Control Tower (SI Control Tower)</code> on first mention so LLM embeddings bind the 32-char canonical name and 16-char shorthand together.<br>
                • <strong>Company Positioning Tagline:</strong> Always spell out <code>ServiceNow is the Super Intelligence Control Tower for business reinvention</code> (75/175 chars ✓) across US and all 13 regional homepages.
              </td>
            </tr>
            <tr>
              <td data-label="Optimization Pillar"><span class="badge b-w3">2. Clean Entity Cutover</span><br><strong>Zero Legacy-Name Eyebrows or Parentheticals</strong></td>
              <td data-label="Technical Mechanism">
                • Including a retired product name inside H1 eyebrows, meta descriptions, or JSON-LD schema causes <strong>entity split and token dilution</strong> in vector retrieval (RAG chunks match both the old and new tokens, causing AI Overviews and chatbots to hedge or cite the outdated name).<br>
                • Instead, <strong>HTTP 301 permanent redirects</strong> (Tab 3) transfer 100% of historical backlink authority and anchor signals at the HTTP layer while keeping on-page DOM text 100% clean.
              </td>
              <td data-label="Exact Implementation Rule">
                • <strong>0 legacy-name eyebrows or parentheticals</strong> anywhere in HTML copy, <code>&lt;meta&gt;</code> tags, or JSON-LD.<br>
                • JSON-LD <code>SoftwareApplication</code> declares <code>"name": "ServiceNow Super Intelligence Control Tower"</code> and <code>"alternateName": ["Super Intelligence Control Tower", "SI Control Tower"]</code>.<br>
                • All {total_redirects_count} legacy URLs and DAM assets return <code>HTTP 301</code> to their new <code>super-intelligence-control-tower</code> paths.
              </td>
            </tr>
            <tr>
              <td data-label="Optimization Pillar"><span class="badge b-warn">3. SERP Pixel &amp; WWSG Fit</span><br><strong>Why <code>SI Control Tower</code> (16 Chars) Is Used on Secondary &amp; Constrained Elements</strong></td>
              <td data-label="Technical Mechanism">
                • Google Desktop SERPs truncate <code>&lt;title&gt;</code> tags at <strong>~600 pixels (~60 characters)</strong> in 20px Roboto, and WWSG 2026 Q3 enforces strict character caps on Arc components (<code>&le;50</code> chars for Accordion/Marquee/Card headings, <code>&le;100</code> chars for Cards/Action items, <code>&le;125</code> chars for <code>alt</code> text).<br>
                • Because <strong><code>SI Control Tower</code> is exactly 16 characters (~138px)</strong> — identical to the old 16-char name — it adds <strong>+0 characters and +0 pixels</strong>!
              </td>
              <td data-label="Exact Implementation Rule">
                • Automatically used on <strong>{total_si_short_saves} instances</strong> across the inventory:<br>
                &nbsp;&nbsp;1. <strong>Compound <code>&lt;title&gt;</code> tags &gt;60 chars</strong> (blogs, webinars, white papers) so <code>" - ServiceNow"</code> never truncates in Google SERPs.<br>
                &nbsp;&nbsp;2. <strong>FAQ Accordion Headings (&le;50 chars)</strong> after the page <code>&lt;h1&gt;</code> has established the full entity.<br>
                &nbsp;&nbsp;3. <strong>Tight Arc UI Cards, Buttons &amp; Secondary Nav (&le;50 / &le;100 chars)</strong>.<br>
                &nbsp;&nbsp;4. <strong>Image <code>alt</code> attributes approaching 125 chars</strong>.<br>
                &nbsp;&nbsp;5. <strong>100% of former <code>AICT</code> acronym mentions</strong>.
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Part 2: Element-by-Element SEO, AIO &amp; WWSG Naming Decision Matrix</h2>
      <p class="desc">Definitive reference for writers, SEO specialists, and AEM engineers showing the exact target form for every DOM &amp; metadata placement.</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>DOM / Metadata Placement</th>
              <th>WWSG / SERP Limit</th>
              <th>Approved Target Form</th>
              <th>Execution Track</th>
              <th>Exact Example &amp; SEO / AIO Rationale</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td data-label="Placement"><strong>1. Flagship Product Page <code>&lt;title&gt;</code></strong><br>{a_link('https://www.servicenow.com/products/ai-control-tower.html')}<br>{a_link('https://www.servicenow.com/uk/products/ai-control-tower.html')}</td>
              <td data-label="Limit">&le; 60 chars<br>(&le; 600px SERP)</td>
              <td data-label="Target Form"><span class="badge b-ok">Super Intelligence Control Tower</span><br><em>or Dual-Keyword</em></td>
              <td data-label="Track"><span class="badge b-ok">🤖 Auto Swap</span></td>
              <td data-label="Example &amp; Rationale">
                • <strong>Primary Target:</strong> <code>Super Intelligence Control Tower - ServiceNow</code> <span class="badge b-ok">[45/60 chars • ~412px ✓]</span><br>
                • <strong>Dual-Keyword SEO Option:</strong> <code>Super Intelligence (SI) Control Tower - ServiceNow</code> <span class="badge b-ok">[50/60 chars • ~458px ✓]</span><br>
                <em>Captures organic search &amp; AI Overview citations for both the full product name and "SI Control Tower".</em>
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>2. Compound <code>&lt;title&gt;</code> Tags</strong><br>(Blogs, Webinars, eBooks, Solution Briefs)</td>
              <td data-label="Limit">&le; 60 chars<br>(&le; 600px SERP)</td>
              <td data-label="Target Form"><span class="badge b-warn">SI Control Tower</span><br>(when Full &gt; 60 chars)</td>
              <td data-label="Track"><span class="badge b-ok">🤖 Auto Swap</span></td>
              <td data-label="Example &amp; Rationale">
                • <strong>Example:</strong> <code>What Is an SI Control Tower? - ServiceNow</code> <span class="badge b-ok">[41/60 chars ✓]</span> (vs. 57 chars with full name).<br>
                • <strong>Rule:</strong> Whenever the 32-char name pushes a compound <code>&lt;title&gt;</code> over 60 chars, the Groovy script automatically uses <code>SI Control Tower</code> (16 chars) to prevent SERP ellipsis truncation.
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>3. Flagship Hero Marquee <code>&lt;h1&gt;</code></strong><br>(<code>&lt;arc-marquee-large slot="heading"&gt;</code>)</td>
              <td data-label="Limit">&le; 50 chars<br>(WWSG Marquee H1)</td>
              <td data-label="Target Form"><span class="badge b-ok">Super Intelligence Control Tower</span></td>
              <td data-label="Track"><span class="badge b-ok">🤖 Auto Swap</span></td>
              <td data-label="Example &amp; Rationale">
                • <strong>Target:</strong> <code>Super Intelligence Control Tower</code> <span class="badge b-ok">[32/50 chars ✓]</span><br>
                • <strong>Clean Cutover:</strong> Direct 1:1 automated swap in the H1 heading with zero legacy eyebrow text.
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>4. Flagship Hero Marquee Lead Copy &amp; <code>&lt;meta description&gt;</code></strong></td>
              <td data-label="Limit">Marquee: &le; 175<br>Meta Desc: &le; 160</td>
              <td data-label="Target Form"><span class="badge b-ok">Super Intelligence Control Tower</span></td>
              <td data-label="Track"><span class="badge b-w1">🛠️ Manual Copy Trim</span><br>(Marquee)<br><span class="badge b-ok">🤖 Auto Swap</span> (Meta)</td>
              <td data-label="Example &amp; Rationale">
                • <strong>Marquee Description (&le;175 chars):</strong> <code>Enforce runtime governance for AI at scale. Super Intelligence Control Tower gives you a central hub to discover, secure, govern, observe, and measure enterprise AI.</code> <span class="badge b-ok">[165/175 chars ✓]</span><br>
                • <strong>SEO Meta Description (&le;160 chars):</strong> <code>Discover, govern, secure, and measure enterprise AI from one central hub with ServiceNow Super Intelligence (SI) Control Tower.</code> <span class="badge b-ok">[127/160 chars ✓]</span>
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>5. Signature Company Positioning Tagline</strong><br>({a_link('https://www.servicenow.com/')}, {a_link('https://www.servicenow.com/uk/')}, Platform Hubs)</td>
              <td data-label="Limit">&le; 175 chars</td>
              <td data-label="Target Form"><span class="badge b-ok">Always Full Name</span></td>
              <td data-label="Track"><span class="badge b-ok">🤖 Auto Swap</span></td>
              <td data-label="Example &amp; Rationale">
                • <strong>Target:</strong> <code>ServiceNow is the Super Intelligence Control Tower for business reinvention</code> <span class="badge b-ok">[75/175 chars ✓]</span><br>
                <em>Essential for corporate brand authority and LLM entity association across US and all 13 regional homepages.</em>
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>6. FAQ Accordion Headings</strong><br>(<code>&lt;arc-accordion-item slot="heading"&gt;</code>)</td>
              <td data-label="Limit">&le; 50 chars<br>(WWSG Accordion Title)</td>
              <td data-label="Target Form"><span class="badge b-warn">SI Control Tower</span><br>(16 chars)</td>
              <td data-label="Track"><span class="badge b-w1">🛠️ Manual Copy Trim</span><br>(Flagship FAQs)</td>
              <td data-label="Example &amp; Rationale">
                • Q1: <code>How is SI Control Tower unique?</code> <span class="badge b-ok">[31/50 chars ✓]</span><br>
                • Q2: <code>Can SI Control Tower discover shadow AI agents?</code> <span class="badge b-ok">[47/50 chars ✓]</span><br>
                • Q3: <code>How does SI Control Tower handle compliance?</code> <span class="badge b-ok">[44/50 chars ✓]</span><br>
                • Q4: <code>Does SI Control Tower work with ServiceNow?</code> <span class="badge b-ok">[43/50 chars ✓]</span><br>
                <em>Because the H1 already establishes <code>Super Intelligence Control Tower</code>, using <code>SI Control Tower</code> in FAQ headings satisfies the 50-char WWSG cap and targets conversational PAA/AIO queries.</em>
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>7. FAQ Accordion Body Answers</strong><br>(<code>&lt;arc-accordion-item&gt; &lt;arc-copy-block&gt;</code>)</td>
              <td data-label="Limit">&le; 250 chars<br>(WWSG Accordion Body)</td>
              <td data-label="Target Form"><span class="badge b-ok">Full Name</span> (A1) +<br><span class="badge b-warn">SI Control Tower</span> (A2–A4)</td>
              <td data-label="Track"><span class="badge b-w1">🛠️ Manual Copy Trim</span></td>
              <td data-label="Example &amp; Rationale">
                • Answer #1 uses <code>Super Intelligence Control Tower</code> (224/250 chars ✓) to reinforce the full entity inside `FAQPage` schema.<br>
                • Answers #3 &amp; #4 use <code>SI Control Tower</code> (210/250 and 177/250 chars ✓) to stay crisp and under the 250-char WWSG cap.
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>8. Secondary Nav, Feature Cards, Action Items &amp; Former <code>AICT</code> Acronyms</strong></td>
              <td data-label="Limit">Nav/Heading: &le; 50<br>Card/Action: &le; 100</td>
              <td data-label="Target Form"><span class="badge b-warn">SI Control Tower</span><br>(when Full overflows)</td>
              <td data-label="Track"><span class="badge b-ok">🤖 Auto Swap</span></td>
              <td data-label="Example &amp; Rationale">
                • Secondary Nav: <code>Super Intelligence Control Tower</code> (32/50 chars ✓ — full name fits!).<br>
                • Action Card (&le;100 chars): <code>Review our product doc to learn what you’ll get from SI Control Tower.</code> <span class="badge b-ok">[70/100 chars ✓]</span>.<br>
                • All former <code>AICT</code> acronyms swap 1:1 to <code>SI Control Tower</code>.
              </td>
            </tr>
            <tr>
              <td data-label="Placement"><strong>9. HTML Image <code>alt</code> Attributes</strong><br>(<code>&lt;arc-image alt="..."&gt;</code>)</td>
              <td data-label="Limit">&le; 125 chars<br>(WWSG Accessibility)</td>
              <td data-label="Target Form"><span class="badge b-warn">SI Control Tower</span><br>(when Full &gt; 125 chars)</td>
              <td data-label="Track"><span class="badge b-ok">🤖 Auto Swap</span></td>
              <td data-label="Example &amp; Rationale">
                • Strips redundant <em>"Video showing / Screenshot of"</em> prefixes and uses <code>SI Control Tower</code> when needed so all 11 flagship <code>alt</code> attributes fit inside 98–103 chars (&le;125 chars ✓).
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Part 3: Ready-to-Embed JSON-LD Structured Data (<code>SoftwareApplication</code> + <code>FAQPage</code> for AIO &amp; Rich Results)</h2>
      <p class="desc">Embed on {a_link('https://www.servicenow.com/products/ai-control-tower.html')} (after move to <code>/products/super-intelligence-control-tower.html</code>) and localized counterparts. Declares both <code>Super Intelligence Control Tower</code> and <code>SI Control Tower</code> with zero legacy product names.</p>
      <pre>&lt;script type="application/ld+json"&gt;
{{
  "@context": "https://schema.org",
  "@graph": [
    {{
      "@type": "SoftwareApplication",
      "@id": "https://www.servicenow.com/products/super-intelligence-control-tower.html#software",
      "name": "ServiceNow Super Intelligence Control Tower",
      "alternateName": [
        "Super Intelligence Control Tower",
        "SI Control Tower"
      ],
      "applicationCategory": "BusinessApplication",
      "ApplicationSubCategory": "Enterprise AI Governance, Observability, and Management Platform",
      "operatingSystem": "ServiceNow AI Platform",
      "url": "https://www.servicenow.com/products/super-intelligence-control-tower.html",
      "description": "Enforce runtime governance for AI at scale. Super Intelligence Control Tower (SI Control Tower) gives you a central hub to discover, secure, govern, observe, and measure enterprise AI.",
      "brand": {{
        "@type": "Brand",
        "name": "ServiceNow"
      }},
      "offers": {{
        "@type": "Offer",
        "url": "https://www.servicenow.com/contact-us-super-intelligence-control-tower.html"
      }}
    }},
    {{
      "@type": "FAQPage",
      "@id": "https://www.servicenow.com/products/super-intelligence-control-tower.html#faq",
      "mainEntity": [
        {{
          "@type": "Question",
          "name": "How is SI Control Tower unique?",
          "acceptedAnswer": {{
            "@type": "Answer",
            "text": "Most vendors only cover one area or their own cloud. Super Intelligence Control Tower unifies AI discovery, observability, governance, security, and value metrics to give you one source of truth for AI assets across clouds and vendors."
          }}
        }},
        {{
          "@type": "Question",
          "name": "Can SI Control Tower discover shadow AI agents?",
          "acceptedAnswer": {{
            "@type": "Answer",
            "text": "Yes. Super Intelligence Control Tower continuously discovers AI agents, models, and workflows across your enterprise—including third-party and shadow AI—and registers them in your unified inventory."
          }}
        }},
        {{
          "@type": "Question",
          "name": "How does SI Control Tower handle compliance?",
          "acceptedAnswer": {{
            "@type": "Answer",
            "text": "SI Control Tower provides automated, auditable compliance gates with out-of-the-box governance workflows and content packs for the EU AI Act, NIST AI RMF, and state regulations like the California AI Transparency Act."
          }}
        }},
        {{
          "@type": "Question",
          "name": "Does SI Control Tower work with ServiceNow?",
          "acceptedAnswer": {{
            "@type": "Answer",
            "text": "Yes. SI Control Tower is built into the ServiceNow AI Platform and your CMDB. Contact your account team to see how it compounds the value of your existing ServiceNow investments."
          }}
        }}
      ]
    }}
  ]
}}
&lt;/script&gt;</pre>
    </div>

    <div class="card">
      <h2>Part 4: Ready-to-Run AEM Groovy Console Script (Executes All {total_auto_instances:,} Automated Swaps Across US + 13 Regional Locales)</h2>
      <p class="desc">Automatically updates JCR text, meta, link, and alt properties across US and all 13 regional sites, using <code>Super Intelligence Control Tower</code> by default and automatically falling back to <code>SI Control Tower</code> when a component property has a strict character cap.</p>
      <pre>// AEM Groovy Console Script — Global + Regional Automated Swap with WWSG Length-Aware "SI Control Tower" Fallback
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

                // 1. Update internal URL &amp; DAM slugs
                if (!val.contains("logo-aicte.svg")) {{
                    val = val.replaceAll("(?i)/products/ai-control-tower", "/products/super-intelligence-control-tower")
                    val = val.replaceAll("(?i)contact-us-aict", "contact-us-super-intelligence-control-tower")
                    val = val.replaceAll("(?i)ebk-aict-discovery-playbook", "ebk-super-intelligence-control-tower-discovery-playbook")
                }}

                // 2. Determine if property has a tight WWSG / SERP character cap
                int maxCap = 0
                if (pName.contains("jcr:title") || pName.contains("navtitle") || pName.contains("heading")) maxCap = 50
                else if (pName.contains("pagetitle") || pName.contains("seotitle")) maxCap = 60
                else if (pName.contains("alt")) maxCap = 125

                // 3. Apply English + 10-language localized replacement (Full Name vs. Short "SI" fallback)
                String fullCandidate = val
                    .replaceAll("(?i)\\\\bAI Control Tower\\\\b", "Super Intelligence Control Tower")
                    .replaceAll("(?i)\\\\bAICT\\\\b", "SI Control Tower")
                    .replaceAll("KI-Kontrollturm", "Super-Intelligenz-Kontrollturm")
                    .replaceAll("(?i)Tour de contr\\\\u00f4le (?:de l'IA|IA)", "Tour de contr\\\\u00f4le Super Intelligence")
                    .replaceAll("(?i)Torre de control de IA", "Torre de control de Superinteligencia")
                    .replaceAll("(?i)Torre de controle de IA", "Torre de controle de Superintelig\\\\u00eancia")
                    .replaceAll("(?i)Torre di controllo (?:dell'IA|AI|IA)", "Torre di controllo della Super Intelligenza")
                    .replaceAll("(?i)AI-controletoren", "Superintelligentie-controletoren")
                    .replaceAll("AI\\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc|AI \\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc", "\\u30b9\\u30fc\\u30d1\\u30fc\\u30a4\\u30f3\\u30c6\\u30ea\\u30b8\\u30a7\\u30f3\\u30b9\\u30fb\\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc")
                    .replaceAll("AI \\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc|AI \\ucee8\\ud2b8\\ub864\\ud0c0\\uc6cc", "\\uc288\\ud37c \\uc778\\ud154\\ub9ac\\uc804\\uc2a4 \\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc")

                if (maxCap &gt; 0 &amp;&amp; fullCandidate.length() &gt; maxCap &amp;&amp; !orig.toLowerCase().contains("business reinvention")) {{
                    val = val
                        .replaceAll("(?i)\\\\bServiceNow AI Control Tower\\\\b", "SI Control Tower")
                        .replaceAll("(?i)\\\\bAI Control Tower\\\\b", "SI Control Tower")
                        .replaceAll("(?i)\\\\bAICT\\\\b", "SI Control Tower")
                        .replaceAll("KI-Kontrollturm", "SI-Kontrollturm")
                        .replaceAll("(?i)Tour de contr\\\\u00f4le (?:de l'IA|IA)", "Tour de contr\\\\u00f4le SI")
                        .replaceAll("(?i)Torre de control de IA", "Torre de control de SI")
                        .replaceAll("(?i)Torre de controle de IA", "Torre de controle de SI")
                        .replaceAll("(?i)Torre di controllo (?:dell'IA|AI|IA)", "Torre di controllo SI")
                        .replaceAll("(?i)AI-controletoren", "SI-controletoren")
                        .replaceAll("AI\\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc|AI \\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc", "SI \\u30b3\\u30f3\\u30c8\\u30ed\\u30fc\\u30eb\\u30bf\\u30ef\\u30fc")
                        .replaceAll("AI \\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc|AI \\ucee8\\ud2b8\\ub864\\ud0c0\\uc6cc", "SI \\ucee8\\ud2b8\\ub864 \\ud0c0\\uc6cc")
                }} else {{
                    val = fullCandidate
                }}

                if (val != orig) {{
                    updatedProps++
                    println "[UPDATE] ${{node.path}}@${{prop.name}} (${{val.length()}} chars)"
                    if (!dryRun) prop.setValue(val)
                }}
            }}
        }}
    }}
}}
if (!dryRun) session.save()
println "Total properties updated: ${{updatedProps}} (dryRun=${{dryRun}})"</pre>
    </div>
  </div>

  <!-- TAB 5: WAVE ROLLOUT & MEDIA / L10N SUMMARY -->
  <div id="tab-waves" class="tab-panel">
    <div class="alert">
      <strong>Wave Rollout Strategy (US Traffic + 13-Locale Regional Inheritance, No Hardcoded Dates):</strong>
      Pages are grouped into <strong>Wave 1, Wave 2, and Wave 3</strong> using the Page Category &amp; Traffic Prioritization model, and <strong>every regional counterpart automatically inherits the Wave of its US parent page</strong>.
    </div>

    <div class="card">
      <h2>Part 1: Wave Rollout Breakdown Across All {len(unique_body_pages)} Pages ({us_pages_total} US + {regional_pages_total} Regional)</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Wave</th>
              <th>Selection Criteria (Page Category + Traffic &amp; Regional Inheritance Rules)</th>
              <th>Pages (US + Regional) &amp; Execution Split</th>
              <th>US YTD Visits (% Share)</th>
              <th>US YTD Conversions (% Share)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td data-label="Wave"><span class="badge b-w1">Wave 1</span></td>
              <td data-label="Selection Criteria">
                • <strong>Global Nav Experience Fragment + Search Prompt JSON:</strong> {len(ssi_only_pages):,} pages auto-updated.<br>
                • <strong>Flagship Product &amp; Playbook Pages (US + All 13 Locales):</strong> {a_link('https://www.servicenow.com/products/ai-control-tower.html')}, {a_link('https://www.servicenow.com/uk/products/ai-control-tower.html')}, etc., plus {a_link('https://www.servicenow.com/contact-us-aict.html')} &amp; {a_link('https://www.servicenow.com/standard/resource-center/ebook/ebk-aict-discovery-playbook.html')}.<br>
                • <strong>Tier-1 Hubs (US + All 13 Regional Locales):</strong> <code>Homepage</code> (including {a_link('https://www.servicenow.com/')} and all 13 regional homepages: {a_link('https://www.servicenow.com/uk/')}, {a_link('https://www.servicenow.com/au/')}, {a_link('https://www.servicenow.com/de/')}, {a_link('https://www.servicenow.com/fr/')}, etc.), <code>Platform - Main Page</code>, <code>Products by Category</code>, <code>What is ServiceNow</code>, <code>Core AI Pages</code>, <code>Events - Homepage</code>, <code>Customer - Homepage</code>, <code>Impact Pages</code>.<br>
                • <strong>High-Traffic / High-Conversion Promoted Pages (+ Regional Counterparts):</strong> Any rebrand page with <strong>YTD Visits &ge; 10,000</strong> OR <strong>YTD Form Conversions &ge; 250</strong> OR <strong>Priority Score &ge; 3.33</strong>.
              </td>
              <td data-label="Pages &amp; Execution Split"><strong>{w1_s['pages']} Pages</strong> ({w1_s['us_pages']} US + {w1_s['reg_pages']} Reg)<br><span style="color:var(--muted);font-size:11px;">({w1_s['instances']:,} total edits)</span><br><span style="color:var(--green);font-size:11.5px;">🤖 {w1_s['auto_pages']} Pure Auto</span> / <span style="color:var(--amber);font-size:11.5px;">🛠️ {w1_s['manual_pages']} Manual Hybrid</span></td>
              <td data-label="US YTD Visits"><strong style="color:var(--green);font-size:15px;">{w1_s['visits']:,}</strong><br><span class="badge b-ok">96.2% of Traffic</span></td>
              <td data-label="US YTD Conversions"><strong style="color:var(--amber);font-size:15px;">{w1_s['conv']:,}</strong><br><span class="badge b-ok">90.8% of Conversions</span></td>
            </tr>
            <tr>
              <td data-label="Wave"><span class="badge b-w2">Wave 2</span></td>
              <td data-label="Selection Criteria">
                • <strong>All Remaining Commercial &amp; Solution Categories (US + All 13 Regional Locales):</strong> <code>Product Pages</code>, <code>Solutions</code>, <code>Capability Pages</code>, <code>Application Pages</code>, <code>Platform - Other Pages</code>, <code>Industry Pages</code>, <code>What is Pages</code>.<br>
                • <strong>Resource Center, Active Events &amp; Corporate Pages (US + Regional):</strong> <code>Resource Center</code>, <code>Ungated ePDF Pages</code>, <code>Events</code>, <code>Contact Us Page</code>, <code>Company Pages</code>.<br>
                • <strong>Active Workflow Editorial Pages:</strong> <code>Workflow</code> articles with <strong>YTD Visits &ge; 100</strong> or <strong>Conversions &ge; 1</strong>.
              </td>
              <td data-label="Pages &amp; Execution Split"><strong>{w2_s['pages']} Pages</strong> ({w2_s['us_pages']} US + {w2_s['reg_pages']} Reg)<br><span style="color:var(--muted);font-size:11px;">({w2_s['instances']:,} total edits)</span><br><span style="color:var(--green);font-size:11.5px;">🤖 {w2_s['auto_pages']} Pure Auto</span> / <span style="color:var(--amber);font-size:11.5px;">🛠️ {w2_s['manual_pages']} Manual Hybrid</span></td>
              <td data-label="US YTD Visits"><strong style="color:var(--amber);font-size:15px;">{w2_s['visits']:,}</strong><br><span class="badge b-w2">3.6% of Traffic</span></td>
              <td data-label="US YTD Conversions"><strong style="color:var(--amber);font-size:15px;">{w2_s['conv']:,}</strong><br><span class="badge b-w2">8.9% of Conversions</span></td>
            </tr>
            <tr>
              <td data-label="Wave"><span class="badge b-w3">Wave 3</span></td>
              <td data-label="Selection Criteria">
                • <strong>Long-Tail Archival Blogs:</strong> 2018–2025 <code>Blogs Pages</code> (only 3,399 combined YTD visits).<br>
                • <strong>Low-Traffic Archival Workflow Articles (&lt;100 YTD visits):</strong> 48 older <code>Workflow</code> articles.<br>
                • <strong>Customer Stories, Partner Finder, University &amp; Docs (US + Regional):</strong> <code>Customer Pages</code>, <code>Partners</code>, <code>University</code>, and <code>Docs</code>.
              </td>
              <td data-label="Pages &amp; Execution Split"><strong>{w3_s['pages']} Pages</strong> ({w3_s['us_pages']} US + {w3_s['reg_pages']} Reg)<br><span style="color:var(--muted);font-size:11px;">({w3_s['instances']:,} total edits)</span><br><span style="color:var(--green);font-size:11.5px;">🤖 {w3_s['auto_pages']} Pure Auto</span> / <span style="color:var(--amber);font-size:11.5px;">🛠️ {w3_s['manual_pages']} Manual Hybrid</span></td>
              <td data-label="US YTD Visits"><strong style="color:var(--blue);font-size:15px;">{w3_s['visits']:,}</strong><br><span class="badge b-w3">0.2% of Traffic</span></td>
              <td data-label="US YTD Conversions"><strong style="color:var(--blue);font-size:15px;">{w3_s['conv']:,}</strong><br><span class="badge b-w3">0.3% of Conversions</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Part 2: Page Category Summary Across All {len(unique_body_pages)} Rebrand Pages (31 Categories)</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Page Category</th>
              <th>Rebrand Pages (US + Regional)</th>
              <th>Automated vs. Manual Hybrid Pages</th>
              <th>US YTD Visits (% Share)</th>
              <th>US Form Conversions</th>
              <th>Wave Split</th>
            </tr>
          </thead>
          <tbody>
            {"".join(cat_summary_rows_html)}
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Part 3: Complete DAM Image &amp; PDF Handoff Summary ({total_baked_images_count} Baked-In Images [{total_baked_designs_count} Visual Designs] + {total_zero_pixel_images_count} Zero-Pixel Slug Images + 24 PDFs)</h2>
      <p class="desc">Comprehensive 4-Layer Discovery (Live 388-Page Component Scan + 8,113 Brand 3.0 Excel DAM Assets + Responsive Variant Probe + Apple Vision OCR): <strong>{total_baked_images_count} images ({total_baked_designs_count} visual designs)</strong> have confirmed baked-in pixel text requiring design work (including <code>welcome-ai-control-tower-.sm.png</code> on <code>/platform/responsible-ai.html</code>), <strong>{total_zero_pixel_images_count} slug-matched images</strong> have 0 baked-in text and only need an automated DAM filename move + 301 redirect, and <strong>13 PDFs</strong> contain live/vector text requiring document updates.</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Execution Mode &amp; Type</th>
              <th>Full Clickable DAM Asset URL &amp; Host Pages</th>
              <th>Specs</th>
              <th>OCR / PDFKit Finding</th>
              <th>Handoff Action</th>
            </tr>
          </thead>
          <tbody>
            {"".join(media_rows_html)}
          </tbody>
        </table>
      </div>
    </div>

    <div class="card">
      <h2>Part 4: Complete 13-Locale Regional Inventory &amp; Localization Handoff ({regional_pages_total} Regional Pages)</h2>
      <p class="desc">Every single regional page with live <code>AI Control Tower</code>, localized equivalents, or <code>/products/ai-control-tower.html</code> links is included in Tabs 1–3 and the CSV.</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Region / Locale</th>
              <th>Pages in Inventory</th>
              <th>Key Live URLs (Homepage &amp; Flagship Product Page)</th>
              <th>Current Live Term</th>
              <th>Target Full Name &amp; Short (<code>SI</code>) Option</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td data-label="Region / Locale"><strong>🇬🇧 United Kingdom (<code>/uk/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['uk']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/uk/')}<br>{a_link('https://www.servicenow.com/uk/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>AI Control Tower</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Super Intelligence Control Tower</strong> (Short: <strong>SI Control Tower</strong>) — <em>🤖 Automated via AEM Live Copy</em></td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇦🇺 Australia (<code>/au/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['au']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/au/')}<br>{a_link('https://www.servicenow.com/au/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>AI Control Tower</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Super Intelligence Control Tower</strong> (Short: <strong>SI Control Tower</strong>) — <em>🤖 Automated via AEM Live Copy</em></td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇮🇳 India (<code>/in/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['in']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/in/')}<br>{a_link('https://www.servicenow.com/in/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>AI Control Tower</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Super Intelligence Control Tower</strong> (Short: <strong>SI Control Tower</strong>) — <em>🤖 Automated via AEM Live Copy</em></td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇩🇪 Germany (<code>/de/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['de']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/de/')}<br>{a_link('https://www.servicenow.com/de/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>KI-Kontrollturm</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Super-Intelligenz-Kontrollturm</strong> (Short: <strong>SI-Kontrollturm</strong>)</td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇫🇷 France (<code>/fr/</code>) &amp; 🇨🇦 Canada FR (<code>/fr-ca/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['fr']} FR + {locale_counts['fr-ca']} FR-CA Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/fr/')}<br>{a_link('https://www.servicenow.com/fr/products/ai-control-tower.html')}<br>{a_link('https://www.servicenow.com/fr-ca/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>Tour de contrôle IA</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Tour de contrôle Super Intelligence</strong> (Short: <strong>Tour de contrôle SI</strong>)</td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇪🇸 Spain (<code>/es/</code>) &amp; 🌎 LATAM (<code>/latam/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['es']} ES + {locale_counts['latam']} LATAM Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/es/')}<br>{a_link('https://www.servicenow.com/es/products/ai-control-tower.html')}<br>{a_link('https://www.servicenow.com/latam/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>Torre de control de IA</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Torre de control de Superinteligencia</strong> (Short: <strong>Torre de control de SI</strong>)</td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇧🇷 Brazil (<code>/br/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['br']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/br/')}<br>{a_link('https://www.servicenow.com/br/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>Torre de controle de IA</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Torre de controle de Superinteligência</strong> (Short: <strong>Torre de controle de SI</strong>)</td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇮🇹 Italy (<code>/it/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['it']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/it/')}<br>{a_link('https://www.servicenow.com/it/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>Torre di controllo AI</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Torre di controllo della Super Intelligenza</strong> (Short: <strong>Torre di controllo SI</strong>)</td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇳🇱 Netherlands (<code>/nl/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['nl']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/nl/')}<br>{a_link('https://www.servicenow.com/nl/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>AI-controletoren</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>Superintelligentie-controletoren</strong> (Short: <strong>SI-controletoren</strong>)</td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇯🇵 Japan (<code>/jp/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['jp']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/jp/')}<br>{a_link('https://www.servicenow.com/jp/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>AI コントロールタワー</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>スーパーインテリジェンス・コントロールタワー</strong> (Short: <strong>SI コントロールタワー</strong>)</td>
            </tr>
            <tr>
              <td data-label="Region / Locale"><strong>🇰🇷 Korea (<code>/kr/</code>)</strong></td>
              <td data-label="Pages in Inventory"><strong>{locale_counts['kr']} Pages</strong></td>
              <td data-label="Key Live URLs" class="mono">{a_link('https://www.servicenow.com/kr/')}<br>{a_link('https://www.servicenow.com/kr/products/ai-control-tower.html')}</td>
              <td data-label="Current Live Term"><code>AI 컨트롤 타워</code></td>
              <td data-label="Target Full &amp; Short Name"><strong>슈퍼 인텔리전스 컨트롤 타워</strong> (Short: <strong>SI 컨트롤 타워</strong>)</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</div>

</body>
</html>
"""

open(OUT_HTML, "w", encoding="utf-8").write(full_html)

OUT_MD = "/Users/tamer.mina/.gemini/antigravity-cli/brain/37667ee2-9f2b-4c38-a3b2-645e77963657/aict_rebrand_plan.md"
plan_md = f"""# Super Intelligence Control Tower (SI Control Tower) — Global & Regional Rebrand Handoff Plan

> [!IMPORTANT]
> **Full Global + Regional Coverage ({len(unique_body_pages)} Pages Across US + 13 Regional Locales + 63 Audited DAM Images)**
> Because the traffic prioritization workbook only tracks US (`en-US`) visits, **zero regional pages are removed due to missing traffic data**. Every regional page that requires a copy or link update is included and inherits the Wave of its US parent page:
> - **{len(unique_body_pages)} Total Body-Match Pages**: **{us_pages_total} US pages** + **{regional_pages_total} Regional pages** across **13 locales** (UK: {locale_counts['uk']}, DE: {locale_counts['de']}, FR: {locale_counts['fr']}, JP: {locale_counts['jp']}, NL: {locale_counts['nl']}, BR: {locale_counts['br']}, AU: {locale_counts['au']}, ES: {locale_counts['es']}, KR: {locale_counts['kr']}, FR-CA: {locale_counts['fr-ca']}, IT: {locale_counts['it']}, LATAM: {locale_counts['latam']}, IN: {locale_counts['in']}).
> - **All 13 Regional Homepages Included in Wave 1**: Includes [servicenow.com/uk/](https://www.servicenow.com/uk/) alongside [servicenow.com/](https://www.servicenow.com/) and all 12 other regional homepages ([/au/](https://www.servicenow.com/au/), [/in/](https://www.servicenow.com/in/), [/de/](https://www.servicenow.com/de/), [/fr/](https://www.servicenow.com/fr/), [/fr-ca/](https://www.servicenow.com/fr-ca/), [/es/](https://www.servicenow.com/es/), [/latam/](https://www.servicenow.com/latam/), [/br/](https://www.servicenow.com/br/), [/it/](https://www.servicenow.com/it/), [/nl/](https://www.servicenow.com/nl/), [/jp/](https://www.servicenow.com/jp/), [/kr/](https://www.servicenow.com/kr/)).
> - **Complete 4-Layer Apple Vision OCR Image Audit (63 Total Images = {total_baked_images_count} Baked-In Pixel Images [{total_baked_designs_count} Visual Designs] + {total_zero_pixel_images_count} Zero-Pixel Slug-Matched Images)**: Combines the 388-page component-proximity scan, the **8,113-row Brand 3.0 DAM Excel export** (`response_results_brand-3.0 assets.xlsx`), responsive `.sm`/`.xs`/`.xl` variant probing, and native macOS Apple Vision `VNRecognizeTextRequest` OCR. Includes [welcome-ai-control-tower-.sm.png](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-backgrounds/welcome-ai-control-tower-.sm.png) (and `.xs.png`) on [servicenow.com/platform/responsible-ai.html](https://www.servicenow.com/platform/responsible-ai.html) (+ 13 regional locales and `/standard/other-documents/responsible-ai-guidelines-2025.html`).
> - **Synchronized Handoff Deliverables**:
>   - **5-Tab Interactive HTML Dashboard**: [aict-rebrand-interactive-report.html](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-interactive-report.html) (`1. 🤖 Automated Updates`, `2. 🛠️ Manual Updates`, `3. 🔀 301 Redirects`, `4. 🔍 SEO & AIO Strategy`, `5. 📊 Wave Rollout & Media/L10n Summary`)
>   - **Master CSV Inventory**: [aict-rebrand-inventory.csv](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-inventory.csv) (with separate `Automated_Edits_Count`, `Manual_Tasks_Count`, `Automated_Changes_Breakdown`, and `Manual_Changes_Breakdown` columns)
>   - **Master Image OCR Inventory JSON**: [master_image_ocr_inventory.json](file:///Users/tamer.mina/Documents/Antigravity/master_image_ocr_inventory.json)
>   - **Generator Script**: [sync_all_three_artifacts.py](file:///Users/tamer.mina/Documents/Antigravity/sync_all_three_artifacts.py)

---

## 1. Automated Updates (Tab 1) vs. Manual Updates (Tab 2) Split

Out of **{len(unique_body_pages)} total body-match pages** (**{total_auto_instances + total_manual_copy_instances:,} text/meta/alt instances**), **{round(pages_100pct_auto / len(unique_body_pages) * 100, 1)}% of pages ({pages_100pct_auto} pages)** are **100% automated**, and **all {len(unique_body_pages)} pages** receive automated DOM/meta/link updates in **Tab 1 (`🤖 Automated Updates`)**. **{pages_needing_manual} hybrid pages** require additional manual tasks in **Tab 2 (`🛠️ Manual Updates`)** — and because those {pages_needing_manual} pages have *both* automated bulk swaps and manual tasks, they appear in **both Tab 1 (showing their automated swaps) and Tab 2 (showing their manual tasks)** with one-click cross-links:

| Track & Owner | Dashboard Tab & Scope | Exact Action & Why It Is Automated or Manual |
| :--- | :--- | :--- |
| **🤖 Track A1: Automated Global Nav XF**<br>*Owner: AEM Platform Admin* | **Tab 1 (`🤖 Automated Updates`)**<br>**2 Centralized XFs** (Cascades to {len(ssi_only_pages):,} pages) | Single Experience Fragment update swaps `"AI Control Tower"` → `"Super Intelligence Control Tower"` (or `"SI Control Tower"`) and updates the nav link across US + 13 locales. |
| **🤖 Track A2: Automated AEM Groovy Swap**<br>*Owner: AEM Engineering* | **Tab 1 (`🤖 Automated Updates`)**<br>**All {len(unique_body_pages)} Pages ({pages_100pct_auto} Pure Auto + {pages_needing_manual} Hybrid)**<br>**{total_auto_instances:,} Automated Instances** | Deterministic AEM Groovy script updates `<title>`, `<meta>`, `<h1>` Hero headings (`Super Intelligence Control Tower`, 32/50 chars ✓), body copy, internal links, and image `alt` tags. Automatically uses `"SI Control Tower"` (16 chars) on **{total_si_short_saves} instances** where the 32-char name would exceed WWSG character limits. |
| **🛠️ Track M1: Manual WWSG Copy Edits**<br>*Owner: Digital UX Copywriter* | **Tab 2 (`🛠️ Manual Updates`)**<br>**{total_manual_copy_instances} Copy Instances** across Flagship Product & Resource pages | Apply the pre-written WWSG trims on [servicenow.com/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) (+ 13 regional copies like [/uk/products/ai-control-tower.html](https://www.servicenow.com/uk/products/ai-control-tower.html) and [/fr/products/ai-control-tower.html](https://www.servicenow.com/fr/products/ai-control-tower.html)) where existing Marquee lead copy or FAQ Accordion answers exceeded 175/250 chars. |
| **🛠️ Track M2: Manual AEM Move & 301 Redirects**<br>*Owner: AEM Authoring & Web Ops* | **Tab 2 (`🛠️ Manual Updates`) & Tab 3 (`🔀 301 Redirects`)**<br>**{total_redirects_count} Total 301 Redirects** ({webpage_redirect_cnt} Webpages + {pdf_redirect_cnt} PDFs + {img_redirect_cnt} Images) | Move AEM page nodes ([/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) and regional/resource counterparts → `.../super-intelligence-control-tower.html`) and deploy the Apache Dispatcher / Akamai `301` redirect ruleset. |
| **🛠️ Track M3: Manual Design & PDF Updates**<br>*Owner: Creative, Docs & Legal* | **Tab 2 (`🛠️ Manual Updates`)**<br>**{total_baked_images_count} Baked-In Images ({total_baked_designs_count} Visual Designs)**<br>**13 Live / Vector PDFs** | Re-render/retouch all **{total_baked_images_count} baked-in pixel images ({total_baked_designs_count} visual designs)** — including [welcome-ai-control-tower-.sm.png](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-backgrounds/welcome-ai-control-tower-.sm.png) on [servicenow.com/platform/responsible-ai.html](https://www.servicenow.com/platform/responsible-ai.html), [itbg-ai-control-tower.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-Cards/itbg-ai-control-tower.sm.jpg), [discover-every-ai-asset.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-screenshots/discover-every-ai-asset.sm.jpg), [govern-and-control-ai-risk.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-screenshots/govern-and-control-ai-risk.sm.jpg), [now-ai-control-tower.sm.png](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/company-library/media/now-ai-control-tower.sm.png), [secure-any-ai-action.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-screenshots/secure-any-ai-action.sm.jpg), etc.; update [ebk-aict-discovery-playbook.pdf](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/resource-center/e-book/ebk-aict-discovery-playbook.pdf), [ai-control-tower-servicenow-subscription-unit-overview.pdf](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/legal/ai-control-tower-servicenow-subscription-unit-overview.pdf), and 11 Resource Center PDFs. |
| **🛠️ Track M4: Manual Localization QA**<br>*Owner: Regional Web Leads* | **Tab 2 (`🛠️ Manual Updates`)**<br>**10 Non-English Flagships**<br>(DE, FR, FR-CA, ES, LATAM, BR, IT, NL, JP, KR) | Linguistic review of gender/article agreement on the 10 translated flagship product pages before publish. |

---

## 2. Dedicated 301 Redirects Inventory (Tab 3 — {total_redirects_count} Total Redirects)

To ensure a **100% clean entity cutover** without legacy product-name eyebrows or parentheticals in on-page copy, **HTTP 301 Permanent Redirects** transfer all historical search authority, backlinks, and bookmarked URLs across **{total_redirects_count} URL and DAM slugs**:
- **{webpage_redirect_cnt} Webpage Slug Redirects**: All 34 crawled body-match slug URLs (including US + 13 regional `/products/ai-control-tower.html`, `/contact-us-aict.html`, `/events/live/ai-control-tower.html`, `/workflow/ai/...`, `/standard/resource-center/...`, `/docs/...`, and regional blogs) + 3 additional US webinar/pricing endpoints (`/products/ai-control-tower/pricing.html`, `/events/on-demand-webinars/put-ai-to-work-with-ai-control-tower.html`, `/lpwbr/meet-ai-control-tower-enterprise-ai-governance-at-machine-speed.html`) + 13 regional `/<locale>/events/live/ai-control-tower.html` endpoints.
- **{pdf_redirect_cnt} DAM PDF Slug Redirects**: All 10 Resource Center eBook/White Paper/Solution Brief PDFs (EN-US, DE-DE, FR-FR) + 1 Legal Subscription Unit Overview PDF (`ai-control-tower-servicenow-subscription-unit-overview.pdf`) + 1 Data Sheet PDF (`ds-ai-control-tower.pdf`), including `.pdf.coredownload` endpoints.
- **{img_redirect_cnt} DAM Image Slug Redirects**: All **{img_redirect_cnt} `ai-control-tower` DAM images** (`.png`, `.jpg`, `.webp` across live pages + the 8,113-asset Brand 3.0 DAM export, including `welcome-ai-control-tower-.sm.png`/`.xs.png`, `itbg-ai-control-tower.sm.jpg`/`.xs.jpg`, `now-ai-control-tower.sm.png`/`.xs.png`, `k26-post-ai-control-tower-xl.png`/`.xs.png`, `homepage-ai-control-tower.sm.jpg`/`.xs.jpg`, `homepage-paitw-ai-control-tower.sm.png`/`.xs.png`, etc.), explicitly excluding false positive `logo-aicte.svg` (All India Council for Technical Education).

---

## 3. SEO & AIO Strategy: `Super Intelligence Control Tower` vs. `SI Control Tower` (Tab 4)

Connected to the **Semrush MCP Server** alongside full live DOM, Google SERP pixel-width (~600px / 60-char `<title>` cap), Schema.org JSON-LD, and LLM Retrieval-Augmented Generation (RAG) entity analysis across all {len(unique_body_pages)} pages:

| Placement | Approved Target Form | Execution Track | SEO, AIO (LLM Citation) & WWSG Rationale |
| :--- | :--- | :--- | :--- |
| **1. Flagship Product `<title>`** ([ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html)) | **`Super Intelligence Control Tower - ServiceNow`** *(45/60 chars, ~412px)*<br>or **`Super Intelligence (SI) Control Tower - ServiceNow`** *(50/60 chars, ~458px)* | **🤖 Auto Swap** | Establishes canonical entity resolution for Google Search, Google AI Overviews, ChatGPT Search, and Perplexity while staying well under the 60-char / 600px SERP truncation limit. |
| **2. Flagship Hero Marquee `<h1>`** | **`Super Intelligence Control Tower`** *(32/50 chars)* | **🤖 Auto Swap** | Clean 1:1 automated swap inside the 50-char WWSG Marquee `<h1>` limit with zero legacy eyebrow text. |
| **3. First Body Mention & `<meta description>`** | **`Super Intelligence Control Tower (SI Control Tower)`** | **🤖 Auto Swap** (Meta)<br>**🛠️ Manual Trim** (Marquee) | Introduces the 16-char shorthand alongside the 32-char canonical name so LLM embeddings and search engines bind both entities together. |
| **4. Company Positioning Tagline** ([servicenow.com/](https://www.servicenow.com/), [servicenow.com/uk/](https://www.servicenow.com/uk/)) | **`ServiceNow is the Super Intelligence Control Tower for business reinvention`** *(75/175 chars)* | **🤖 Auto Swap** | Always spell out the full 32-char name in the signature company positioning tagline across US and all 13 regional homepages. |
| **5. Compound `<title>` Tags >60 Chars** (Blogs, Webinars, White Papers) | **`SI Control Tower`** *(16 chars, +0 char / +0px delta)* | **🤖 Auto Swap** | Prevents Google SERP truncation of `" - ServiceNow"` on long resource/blog titles (e.g., `What Is an SI Control Tower? - ServiceNow`, 41/60 chars ✓). |
| **6. Constrained UI Components & Former `AICT` Mentions** | **`SI Control Tower`** *(16 chars, +0 char / +0px delta)* | **🤖 Auto Swap** | Automatically applied on **{total_si_short_saves} instances** where +16 chars would overflow a 50-char FAQ Accordion heading, 100-char card/action link, or 125-char image `alt` attribute, and replaces all `AICT` acronyms 1:1. |
| **7. JSON-LD `SoftwareApplication` Schema** | `name`: **`ServiceNow Super Intelligence Control Tower`**<br>`alternateName`: **`["Super Intelligence Control Tower", "SI Control Tower"]`** | **🤖 Auto Swap** | Declares both the canonical name and official shorthand in structured data with zero legacy product names. |

---

## 4. Regional Homepage & 13-Locale Coverage (Including UK)

All **{regional_pages_total} regional pages across 13 locales** are mapped to their US parent pages so no regional market is left with legacy copy or broken links:

| Locale & Pages | Clickable Regional Homepage & Flagship Product Page | Target Full Name / Short `SI` Form |
| :--- | :--- | :--- |
| **🇺🇸 US** (`en-US`) — **{locale_counts['us']} Pages** | • Homepage: [servicenow.com/](https://www.servicenow.com/)<br>• Flagship: [/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇬🇧 UK** (`en-GB`) — **{locale_counts['uk']} Pages** | • Homepage: [servicenow.com/uk/](https://www.servicenow.com/uk/)<br>• Flagship: [/uk/products/ai-control-tower.html](https://www.servicenow.com/uk/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇦🇺 Australia** (`en-AU`) — **{locale_counts['au']} Pages** | • Homepage: [servicenow.com/au/](https://www.servicenow.com/au/)<br>• Flagship: [/au/products/ai-control-tower.html](https://www.servicenow.com/au/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇮🇳 India** (`en-IN`) — **{locale_counts['in']} Pages** | • Homepage: [servicenow.com/in/](https://www.servicenow.com/in/)<br>• Flagship: [/in/products/ai-control-tower.html](https://www.servicenow.com/in/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇩🇪 Germany** (`de-DE`) — **{locale_counts['de']} Pages** | • Homepage: [servicenow.com/de/](https://www.servicenow.com/de/)<br>• Flagship: [/de/products/ai-control-tower.html](https://www.servicenow.com/de/products/ai-control-tower.html) | **Super-Intelligenz-Kontrollturm**<br>Short: **SI-Kontrollturm** |
| **🇫🇷 France** (`fr-FR`) — **{locale_counts['fr']} Pages** | • Homepage: [servicenow.com/fr/](https://www.servicenow.com/fr/)<br>• Flagship: [/fr/products/ai-control-tower.html](https://www.servicenow.com/fr/products/ai-control-tower.html) | **Tour de contrôle Super Intelligence**<br>Short: **Tour de contrôle SI** |
| **🇨🇦 Canada FR** (`fr-CA`) — **{locale_counts['fr-ca']} Pages** | • Homepage: [servicenow.com/fr-ca/](https://www.servicenow.com/fr-ca/)<br>• Flagship: [/fr-ca/products/ai-control-tower.html](https://www.servicenow.com/fr-ca/products/ai-control-tower.html) | **Tour de contrôle Super Intelligence**<br>Short: **Tour de contrôle SI** |
| **🇪🇸 Spain** (`es-ES`) — **{locale_counts['es']} Pages** | • Homepage: [servicenow.com/es/](https://www.servicenow.com/es/)<br>• Flagship: [/es/products/ai-control-tower.html](https://www.servicenow.com/es/products/ai-control-tower.html) | **Torre de control de Superinteligencia**<br>Short: **Torre de control de SI** |
| **🌎 LATAM** (`es-419`) — **{locale_counts['latam']} Pages** | • Homepage: [servicenow.com/latam/](https://www.servicenow.com/latam/)<br>• Flagship: [/latam/products/ai-control-tower.html](https://www.servicenow.com/latam/products/ai-control-tower.html) | **Torre de control de Superinteligencia**<br>Short: **Torre de control de SI** |
| **🇧🇷 Brazil** (`pt-BR`) — **{locale_counts['br']} Pages** | • Homepage: [servicenow.com/br/](https://www.servicenow.com/br/)<br>• Flagship: [/br/products/ai-control-tower.html](https://www.servicenow.com/br/products/ai-control-tower.html) | **Torre de controle de Superinteligência**<br>Short: **Torre de controle de SI** |
| **🇮🇹 Italy** (`it-IT`) — **{locale_counts['it']} Pages** | • Homepage: [servicenow.com/it/](https://www.servicenow.com/it/)<br>• Flagship: [/it/products/ai-control-tower.html](https://www.servicenow.com/it/products/ai-control-tower.html) | **Torre di controllo della Super Intelligenza**<br>Short: **Torre di controllo SI** |
| **🇳🇱 Netherlands** (`nl-NL`) — **{locale_counts['nl']} Pages** | • Homepage: [servicenow.com/nl/](https://www.servicenow.com/nl/)<br>• Flagship: [/nl/products/ai-control-tower.html](https://www.servicenow.com/nl/products/ai-control-tower.html) | **Superintelligentie-controletoren**<br>Short: **SI-controletoren** |
| **🇯🇵 Japan** (`ja-JP`) — **{locale_counts['jp']} Pages** | • Homepage: [servicenow.com/jp/](https://www.servicenow.com/jp/)<br>• Flagship: [/jp/products/ai-control-tower.html](https://www.servicenow.com/jp/products/ai-control-tower.html) | **スーパーインテリジェンス・コントロールタワー**<br>Short: **SI コントロールタワー** |
| **🇰🇷 Korea** (`ko-KR`) — **{locale_counts['kr']} Pages** | • Homepage: [servicenow.com/kr/](https://www.servicenow.com/kr/)<br>• Flagship: [/kr/products/ai-control-tower.html](https://www.servicenow.com/kr/products/ai-control-tower.html) | **슈퍼 인텔리전스 컨트롤 타워**<br>Short: **SI 컨트롤 타워** |

---

## 5. Wave Rollout Summary (Tab 5 — No Hardcoded Dates)

Regional pages automatically inherit the Wave assignment of their US (`en-US`) parent page:

| Wave | Pages & Regional Split | Execution Split & Traffic | Included Page Groups |
| :--- | :--- | :--- | :--- |
| **Wave 1** | **{w1_s['pages']} Total Pages**<br>• 🇺🇸 {w1_s['us_pages']} US Pages<br>• 🌍 {w1_s['reg_pages']} Regional Pages<br>*({w1_s['instances']:,} total edits)* | • **🤖 {w1_s['auto_pages']} Pure Auto / 🛠️ {w1_s['manual_pages']} Hybrid Manual**<br>• **{w1_s['visits']:,} US Visits** *(96.2%)*<br>• **{w1_s['conv']:,} Conversions** *(90.8%)* | Global Nav/Search XFs + US & all 13 Regional Homepages ([servicenow.com/](https://www.servicenow.com/), [servicenow.com/uk/](https://www.servicenow.com/uk/), [/de/](https://www.servicenow.com/de/), [/fr/](https://www.servicenow.com/fr/), etc.) + Flagship Product Pages ([/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) & all 13 locales) + Top Traffic Hub Pages ([/ai.html](https://www.servicenow.com/ai.html), [/now-platform.html](https://www.servicenow.com/now-platform.html), etc.). |
| **Wave 2** | **{w2_s['pages']} Total Pages**<br>• 🇺🇸 {w2_s['us_pages']} US Pages<br>• 🌍 {w2_s['reg_pages']} Regional Pages<br>*({w2_s['instances']:,} total edits)* | • **🤖 {w2_s['auto_pages']} Pure Auto / 🛠️ {w2_s['manual_pages']} Hybrid Manual**<br>• **{w2_s['visits']:,} US Visits** *(3.6%)*<br>• **{w2_s['conv']:,} Conversions** *(8.9%)* | All remaining Commercial Product, Solution, Capability, Application, Platform, Industry, Event, and Resource Center pages across US + 13 regional locales. |
| **Wave 3** | **{w3_s['pages']} Total Pages**<br>• 🇺🇸 {w3_s['us_pages']} US Pages<br>• 🌍 {w3_s['reg_pages']} Regional Pages<br>*({w3_s['instances']:,} total edits)* | • **🤖 {w3_s['auto_pages']} Pure Auto / 🛠️ {w3_s['manual_pages']} Hybrid Manual**<br>• **{w3_s['visits']:,} US Visits** *(0.2%)*<br>• **{w3_s['conv']:,} Conversions** *(0.3%)* | Long-tail Archival Blogs, Low-traffic Workflow articles (<100 visits), Customer Stories, Partners, University, and Docs across US + 13 regional locales. |
| **Total** | **{len(unique_body_pages)} Total Pages**<br>• **🇺🇸 {us_pages_total} US**<br>• **🌍 {regional_pages_total} Regional** | • **🤖 {pages_100pct_auto} Pure Auto (94.9%)**<br>• **🛠️ {pages_needing_manual} Hybrid Manual (5.1%)**<br>• **{total_rebrand_visits:,} US Visits (100%)** | Complete global + regional rebrand handoff inventory. |
"""

open(OUT_MD, "w", encoding="utf-8").write(plan_md)
print(f"Successfully generated 5-tab handoff dashboard {OUT_HTML}, {OUT_CSV}, and {OUT_MD}")
print(f"Total Body-Match Pages: {len(unique_body_pages)} (US: {us_pages_total} + Regional: {regional_pages_total})")
print(f"Pages 100% Auto: {pages_100pct_auto} | Pages Needing Manual: {pages_needing_manual} | Auto Instances: {total_auto_instances} | Manual Copy Instances: {total_manual_copy_instances} | Manual Tasks All: {total_manual_tasks_all} | SI Short Fits: {total_si_short_saves}")
print(f"Total 301 Redirects: {total_redirects_count} (Webpages: {webpage_redirect_cnt}, PDFs: {pdf_redirect_cnt}, Images: {img_redirect_cnt})")

