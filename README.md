# Super Intelligence Control Tower — Global & Regional Rebrand Handoff Plan

> [!IMPORTANT]
> **Product Marketing Naming & Automation Direction Applied Across All 933 Pages (US + 13 Regional Locales + 63 Audited DAM Images)**
> - **Standard Product Name**: **`Super Intelligence Control Tower`** (32 chars). We do **not** use `SI Control Tower` anywhere—all copy, headings, metadata, and former `AICT` acronyms transition to `Super Intelligence Control Tower`.
> - **CTA Button Name**: **`Control Tower`** (13 chars) on buttons (e.g., `See Control Tower` [17/25 chars], `Explore Control Tower` [21/25 chars], `Extend Control Tower` [20/25 chars], and localized button equivalents).
> - **Awareness & First-Mention CTA Context Bridge Standard (`Manual: Context Bridge`)**: Across **128 pages** (all awareness-categorized pages—`Homepage`, `What is ServiceNow`, `Platform - Main Page`, `Impact Pages`, `Events - AI Summits Hub Page`, and tagline-only `Solutions` pages—plus first-mention `Control Tower` CTA button pages across US + 13 locales), a **Manual Context Bridge** copy update is queued to introduce `Super Intelligence Control Tower` in the surrounding Marquee/Feature paragraph (e.g., Homepage Marquee: *"Most companies have AI but few have the data to ground it, the workflows to deploy it, and the security to govern it. Super Intelligence Control Tower is built for all of it. So AI doesn't just think, it acts."*).
> - **All On-Page Image Swaps Are Automated (`Auto`)**: Once updated image copies are uploaded to the new DAM paths (`super-intelligence-control-tower`), the AEM Groovy automation script swaps out the old image paths and puts in the new image paths across all webpages automatically.
> - **WWSG Body Copy Edits Skipped**: Body copy, Marquee lead descriptions, and FAQ answers are 100% automated direct swaps (`AI Control Tower` -> `Super Intelligence Control Tower`) to minimize manual work.
> - **Oversized Headings & Titles Flagged for Manual Decision**: **82 instances across 45 pages** are flagged as `Manual: Name Too Long` where `Super Intelligence Control Tower` causes a constrained heading (`<h2>`/`<h3>`/`<h4>`, card heading, FAQ question title, or tab label >50 chars) or compound `<title>` (>60 chars) to exceed its character limit.
> - **Interactive Dashboard**: [aict-rebrand-interactive-report.html](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-interactive-report.html)
> - **Master CSV Inventory**: [aict-rebrand-inventory.csv](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-inventory.csv)

---

## 1. Automated vs. Manual Execution Split

| Track & Owner | Scope | Action Summary |
| :--- | :--- | :--- |
| **Track A1: Automated Global Nav XF**<br>*Owner: AEM Platform Admin* | **2 Centralized XFs** (4,230 pages) | Single Experience Fragment update swaps `"AI Control Tower"` to `"Super Intelligence Control Tower"` and updates the nav link across US + 13 locales. |
| **Track A2: Automated AEM Groovy Swap**<br>*Owner: AEM Engineering* | **All 933 Pages** (754 Pure Auto + 179 Hybrid)<br>**2,917 Automated Swaps** | Updates `<title>`, `<meta>`, `<h1>` Hero headings (`Super Intelligence Control Tower`, 32/50 chars), all body copy & FAQ answers, CTA buttons (`See Control Tower`, `Explore Control Tower`, `Extend Control Tower`), company taglines (`ServiceNow is the Super Intelligence Control Tower for business reinvention`), internal links, **all on-page DAM image paths (`old image -> new image copy`)**, and image `alt` tags. |
| **Track M0: Awareness & CTA Context Bridge Copy**<br>*Owner: Digital UX Copywriter / PMM* | **128 Pages** (US + 13 locales) | Update Marquee/Feature lead paragraphs on awareness-categorized pages (`Homepage`, `Platform`, `What is ServiceNow`, `Impact`, `AI Summits`, `Solutions`) and first-mention `Control Tower` CTA button pages so `Super Intelligence Control Tower` is explicitly introduced in context. |
| **Track M1: Flagged Oversized Headings & `<title>` Tags**<br>*Owner: Digital UX Copywriter / PMM* | **82 Flagged Instances** across **45 pages** | Review constrained headings (>50 chars) and compound `<title>` tags (>60 chars) where `"Super Intelligence Control Tower"` is too long and decide the replacement text. |
| **Track M2: Manual AEM Move & 301 Redirects**<br>*Owner: AEM Authoring & Web Ops* | **125 Total 301 Redirects** (50 Webpages + 12 PDFs + 63 Images) | Move AEM page nodes (`/products/ai-control-tower.html` -> `/products/super-intelligence-control-tower.html`) and deploy Apache Dispatcher / Akamai 301 rules. |
| **Track M3: DAM Graphic Copy Upload & PDF Updates**<br>*Owner: Creative, Docs & Legal* | **35 Baked-In Image Copies (19 Visual Designs)** + **13 PDFs** | Upload updated graphic copies at new `super-intelligence-control-tower` DAM filenames (page swaps are 100% automated) and update 13 live/vector PDFs. |
| **Track M4: Manual Localization QA**<br>*Owner: Regional Web Leads* | **10 Non-English Flagship Pages** | Linguistic review of gender/article agreement on DE, FR, FR-CA, ES, LATAM, BR, IT, NL, JP, and KR flagship product pages. |
