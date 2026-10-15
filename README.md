# ServiceNow Rebrand Handoff: AI Control Tower → Super Intelligence (SI) Control Tower

- **🌐 Live Interactive Handoff Dashboard (GitHub Pages):** [https://tamer-mina.github.io/servicenow-si-control-tower-rebrand/](https://tamer-mina.github.io/servicenow-si-control-tower-rebrand/)
- **📊 Interactive HTML Report File:** [`aict-rebrand-interactive-report.html`](./aict-rebrand-interactive-report.html)
- **📥 Master CSV Inventory:** [`aict-rebrand-inventory.csv`](./aict-rebrand-inventory.csv)
- **🖼️ Master Image OCR Inventory JSON (63 Images: 35 Baked-In + 28 Zero-Pixel):** [`master_image_ocr_inventory.json`](./master_image_ocr_inventory.json)
- **📋 Executive Handoff Plan (Markdown):** [`aict_rebrand_plan.md`](./aict_rebrand_plan.md)

---

# Super Intelligence Control Tower (SI Control Tower) — Global & Regional Rebrand Handoff Plan

> [!IMPORTANT]
> **Full Global + Regional Coverage (933 Pages Across US + 13 Regional Locales + 63 Audited DAM Images)**
> Because the traffic prioritization workbook only tracks US (`en-US`) visits, **zero regional pages are removed due to missing traffic data**. Every regional page that requires a copy or link update is included and inherits the Wave of its US parent page:
> - **933 Total Body-Match Pages**: **363 US pages** + **570 Regional pages** across **13 locales** (UK: 46, DE: 50, FR: 48, JP: 46, NL: 45, BR: 44, AU: 43, ES: 43, KR: 43, FR-CA: 42, IT: 41, LATAM: 41, IN: 38).
> - **All 13 Regional Homepages Included in Wave 1**: Includes [servicenow.com/uk/](https://www.servicenow.com/uk/) alongside [servicenow.com/](https://www.servicenow.com/) and all 12 other regional homepages ([/au/](https://www.servicenow.com/au/), [/in/](https://www.servicenow.com/in/), [/de/](https://www.servicenow.com/de/), [/fr/](https://www.servicenow.com/fr/), [/fr-ca/](https://www.servicenow.com/fr-ca/), [/es/](https://www.servicenow.com/es/), [/latam/](https://www.servicenow.com/latam/), [/br/](https://www.servicenow.com/br/), [/it/](https://www.servicenow.com/it/), [/nl/](https://www.servicenow.com/nl/), [/jp/](https://www.servicenow.com/jp/), [/kr/](https://www.servicenow.com/kr/)).
> - **Complete 4-Layer Apple Vision OCR Image Audit (63 Total Images = 35 Baked-In Pixel Images [19 Visual Designs] + 28 Zero-Pixel Slug-Matched Images)**: Combines the 388-page component-proximity scan, the **8,113-row Brand 3.0 DAM Excel export** (`response_results_brand-3.0 assets.xlsx`), responsive `.sm`/`.xs`/`.xl` variant probing, and native macOS Apple Vision `VNRecognizeTextRequest` OCR. Includes [welcome-ai-control-tower-.sm.png](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-backgrounds/welcome-ai-control-tower-.sm.png) (and `.xs.png`) on [servicenow.com/platform/responsible-ai.html](https://www.servicenow.com/platform/responsible-ai.html) (+ 13 regional locales and `/standard/other-documents/responsible-ai-guidelines-2025.html`).
> - **Synchronized Handoff Deliverables**:
>   - **5-Tab Interactive HTML Dashboard**: [aict-rebrand-interactive-report.html](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-interactive-report.html) (`1. 🤖 Automated Updates`, `2. 🛠️ Manual Updates`, `3. 🔀 301 Redirects`, `4. 🔍 SEO & AIO Strategy`, `5. 📊 Wave Rollout & Media/L10n Summary`)
>   - **Master CSV Inventory**: [aict-rebrand-inventory.csv](file:///Users/tamer.mina/Documents/Antigravity/aict-rebrand-inventory.csv) (with separate `Automated_Edits_Count`, `Manual_Tasks_Count`, `Automated_Changes_Breakdown`, and `Manual_Changes_Breakdown` columns)
>   - **Master Image OCR Inventory JSON**: [master_image_ocr_inventory.json](file:///Users/tamer.mina/Documents/Antigravity/master_image_ocr_inventory.json)
>   - **Generator Script**: [sync_all_three_artifacts.py](file:///Users/tamer.mina/Documents/Antigravity/sync_all_three_artifacts.py)

---

## 1. Automated Updates (Tab 1) vs. Manual Updates (Tab 2) Split

Out of **933 total body-match pages** (**3,040 text/meta/alt instances**), **80.4% of pages (750 pages)** are **100% automated**, and **all 933 pages** receive automated DOM/meta/link updates in **Tab 1 (`🤖 Automated Updates`)**. **183 hybrid pages** require additional manual tasks in **Tab 2 (`🛠️ Manual Updates`)** — and because those 183 pages have *both* automated bulk swaps and manual tasks, they appear in **both Tab 1 (showing their automated swaps) and Tab 2 (showing their manual tasks)** with one-click cross-links:

| Track & Owner | Dashboard Tab & Scope | Exact Action & Why It Is Automated or Manual |
| :--- | :--- | :--- |
| **🤖 Track A1: Automated Global Nav XF**<br>*Owner: AEM Platform Admin* | **Tab 1 (`🤖 Automated Updates`)**<br>**2 Centralized XFs** (Cascades to 4,230 pages) | Single Experience Fragment update swaps `"AI Control Tower"` → `"Super Intelligence Control Tower"` (or `"SI Control Tower"`) and updates the nav link across US + 13 locales. |
| **🤖 Track A2: Automated AEM Groovy Swap**<br>*Owner: AEM Engineering* | **Tab 1 (`🤖 Automated Updates`)**<br>**All 933 Pages (750 Pure Auto + 183 Hybrid)**<br>**2,921 Automated Instances** | Deterministic AEM Groovy script updates `<title>`, `<meta>`, `<h1>` Hero headings (`Super Intelligence Control Tower`, 32/50 chars ✓), body copy, internal links, and image `alt` tags. Automatically uses `"SI Control Tower"` (16 chars) on **634 instances** where the 32-char name would exceed WWSG character limits. |
| **🛠️ Track M1: Manual WWSG Copy Edits**<br>*Owner: Digital UX Copywriter* | **Tab 2 (`🛠️ Manual Updates`)**<br>**119 Copy Instances** across Flagship Product & Resource pages | Apply the pre-written WWSG trims on [servicenow.com/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) (+ 13 regional copies like [/uk/products/ai-control-tower.html](https://www.servicenow.com/uk/products/ai-control-tower.html) and [/fr/products/ai-control-tower.html](https://www.servicenow.com/fr/products/ai-control-tower.html)) where existing Marquee lead copy or FAQ Accordion answers exceeded 175/250 chars. |
| **🛠️ Track M2: Manual AEM Move & 301 Redirects**<br>*Owner: AEM Authoring & Web Ops* | **Tab 2 (`🛠️ Manual Updates`) & Tab 3 (`🔀 301 Redirects`)**<br>**105 Total 301 Redirects** (50 Webpages + 12 PDFs + 43 Images) | Move AEM page nodes ([/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) and regional/resource counterparts → `.../super-intelligence-control-tower.html`) and deploy the Apache Dispatcher / Akamai `301` redirect ruleset. |
| **🛠️ Track M3: Manual Design & PDF Updates**<br>*Owner: Creative, Docs & Legal* | **Tab 2 (`🛠️ Manual Updates`)**<br>**35 Baked-In Images (19 Visual Designs)**<br>**13 Live / Vector PDFs** | Re-render/retouch all **35 baked-in pixel images (19 visual designs)** — including [welcome-ai-control-tower-.sm.png](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-backgrounds/welcome-ai-control-tower-.sm.png) on [servicenow.com/platform/responsible-ai.html](https://www.servicenow.com/platform/responsible-ai.html), [itbg-ai-control-tower.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-Cards/itbg-ai-control-tower.sm.jpg), [discover-every-ai-asset.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-screenshots/discover-every-ai-asset.sm.jpg), [govern-and-control-ai-risk.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-screenshots/govern-and-control-ai-risk.sm.jpg), [now-ai-control-tower.sm.png](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/company-library/media/now-ai-control-tower.sm.png), [secure-any-ai-action.sm.jpg](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/images/ds-screenshots/secure-any-ai-action.sm.jpg), etc.; update [ebk-aict-discovery-playbook.pdf](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/resource-center/e-book/ebk-aict-discovery-playbook.pdf), [ai-control-tower-servicenow-subscription-unit-overview.pdf](https://www.servicenow.com/content/dam/servicenow-assets/public/en-us/doc-type/legal/ai-control-tower-servicenow-subscription-unit-overview.pdf), and 11 Resource Center PDFs. |
| **🛠️ Track M4: Manual Localization QA**<br>*Owner: Regional Web Leads* | **Tab 2 (`🛠️ Manual Updates`)**<br>**10 Non-English Flagships**<br>(DE, FR, FR-CA, ES, LATAM, BR, IT, NL, JP, KR) | Linguistic review of gender/article agreement on the 10 translated flagship product pages before publish. |

---

## 2. Dedicated 301 Redirects Inventory (Tab 3 — 105 Total Redirects)

To ensure a **100% clean entity cutover** without legacy product-name eyebrows or parentheticals in on-page copy, **HTTP 301 Permanent Redirects** transfer all historical search authority, backlinks, and bookmarked URLs across **105 URL and DAM slugs**:
- **50 Webpage Slug Redirects**: All 34 crawled body-match slug URLs (including US + 13 regional `/products/ai-control-tower.html`, `/contact-us-aict.html`, `/events/live/ai-control-tower.html`, `/workflow/ai/...`, `/standard/resource-center/...`, `/docs/...`, and regional blogs) + 3 additional US webinar/pricing endpoints (`/products/ai-control-tower/pricing.html`, `/events/on-demand-webinars/put-ai-to-work-with-ai-control-tower.html`, `/lpwbr/meet-ai-control-tower-enterprise-ai-governance-at-machine-speed.html`) + 13 regional `/<locale>/events/live/ai-control-tower.html` endpoints.
- **12 DAM PDF Slug Redirects**: All 10 Resource Center eBook/White Paper/Solution Brief PDFs (EN-US, DE-DE, FR-FR) + 1 Legal Subscription Unit Overview PDF (`ai-control-tower-servicenow-subscription-unit-overview.pdf`) + 1 Data Sheet PDF (`ds-ai-control-tower.pdf`), including `.pdf.coredownload` endpoints.
- **43 DAM Image Slug Redirects**: All **43 `ai-control-tower` DAM images** (`.png`, `.jpg`, `.webp` across live pages + the 8,113-asset Brand 3.0 DAM export, including `welcome-ai-control-tower-.sm.png`/`.xs.png`, `itbg-ai-control-tower.sm.jpg`/`.xs.jpg`, `now-ai-control-tower.sm.png`/`.xs.png`, `k26-post-ai-control-tower-xl.png`/`.xs.png`, `homepage-ai-control-tower.sm.jpg`/`.xs.jpg`, `homepage-paitw-ai-control-tower.sm.png`/`.xs.png`, etc.), explicitly excluding false positive `logo-aicte.svg` (All India Council for Technical Education).

---

## 3. SEO & AIO Strategy: `Super Intelligence Control Tower` vs. `SI Control Tower` (Tab 4)

Connected to the **Semrush MCP Server** alongside full live DOM, Google SERP pixel-width (~600px / 60-char `<title>` cap), Schema.org JSON-LD, and LLM Retrieval-Augmented Generation (RAG) entity analysis across all 933 pages:

| Placement | Approved Target Form | Execution Track | SEO, AIO (LLM Citation) & WWSG Rationale |
| :--- | :--- | :--- | :--- |
| **1. Flagship Product `<title>`** ([ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html)) | **`Super Intelligence Control Tower - ServiceNow`** *(45/60 chars, ~412px)*<br>or **`Super Intelligence (SI) Control Tower - ServiceNow`** *(50/60 chars, ~458px)* | **🤖 Auto Swap** | Establishes canonical entity resolution for Google Search, Google AI Overviews, ChatGPT Search, and Perplexity while staying well under the 60-char / 600px SERP truncation limit. |
| **2. Flagship Hero Marquee `<h1>`** | **`Super Intelligence Control Tower`** *(32/50 chars)* | **🤖 Auto Swap** | Clean 1:1 automated swap inside the 50-char WWSG Marquee `<h1>` limit with zero legacy eyebrow text. |
| **3. First Body Mention & `<meta description>`** | **`Super Intelligence Control Tower (SI Control Tower)`** | **🤖 Auto Swap** (Meta)<br>**🛠️ Manual Trim** (Marquee) | Introduces the 16-char shorthand alongside the 32-char canonical name so LLM embeddings and search engines bind both entities together. |
| **4. Company Positioning Tagline** ([servicenow.com/](https://www.servicenow.com/), [servicenow.com/uk/](https://www.servicenow.com/uk/)) | **`ServiceNow is the Super Intelligence Control Tower for business reinvention`** *(75/175 chars)* | **🤖 Auto Swap** | Always spell out the full 32-char name in the signature company positioning tagline across US and all 13 regional homepages. |
| **5. Compound `<title>` Tags >60 Chars** (Blogs, Webinars, White Papers) | **`SI Control Tower`** *(16 chars, +0 char / +0px delta)* | **🤖 Auto Swap** | Prevents Google SERP truncation of `" - ServiceNow"` on long resource/blog titles (e.g., `What Is an SI Control Tower? - ServiceNow`, 41/60 chars ✓). |
| **6. Constrained UI Components & Former `AICT` Mentions** | **`SI Control Tower`** *(16 chars, +0 char / +0px delta)* | **🤖 Auto Swap** | Automatically applied on **634 instances** where +16 chars would overflow a 50-char FAQ Accordion heading, 100-char card/action link, or 125-char image `alt` attribute, and replaces all `AICT` acronyms 1:1. |
| **7. JSON-LD `SoftwareApplication` Schema** | `name`: **`ServiceNow Super Intelligence Control Tower`**<br>`alternateName`: **`["Super Intelligence Control Tower", "SI Control Tower"]`** | **🤖 Auto Swap** | Declares both the canonical name and official shorthand in structured data with zero legacy product names. |

---

## 4. Regional Homepage & 13-Locale Coverage (Including UK)

All **570 regional pages across 13 locales** are mapped to their US parent pages so no regional market is left with legacy copy or broken links:

| Locale & Pages | Clickable Regional Homepage & Flagship Product Page | Target Full Name / Short `SI` Form |
| :--- | :--- | :--- |
| **🇺🇸 US** (`en-US`) — **363 Pages** | • Homepage: [servicenow.com/](https://www.servicenow.com/)<br>• Flagship: [/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇬🇧 UK** (`en-GB`) — **46 Pages** | • Homepage: [servicenow.com/uk/](https://www.servicenow.com/uk/)<br>• Flagship: [/uk/products/ai-control-tower.html](https://www.servicenow.com/uk/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇦🇺 Australia** (`en-AU`) — **43 Pages** | • Homepage: [servicenow.com/au/](https://www.servicenow.com/au/)<br>• Flagship: [/au/products/ai-control-tower.html](https://www.servicenow.com/au/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇮🇳 India** (`en-IN`) — **38 Pages** | • Homepage: [servicenow.com/in/](https://www.servicenow.com/in/)<br>• Flagship: [/in/products/ai-control-tower.html](https://www.servicenow.com/in/products/ai-control-tower.html) | **Super Intelligence Control Tower**<br>Short: **SI Control Tower** |
| **🇩🇪 Germany** (`de-DE`) — **50 Pages** | • Homepage: [servicenow.com/de/](https://www.servicenow.com/de/)<br>• Flagship: [/de/products/ai-control-tower.html](https://www.servicenow.com/de/products/ai-control-tower.html) | **Super-Intelligenz-Kontrollturm**<br>Short: **SI-Kontrollturm** |
| **🇫🇷 France** (`fr-FR`) — **48 Pages** | • Homepage: [servicenow.com/fr/](https://www.servicenow.com/fr/)<br>• Flagship: [/fr/products/ai-control-tower.html](https://www.servicenow.com/fr/products/ai-control-tower.html) | **Tour de contrôle Super Intelligence**<br>Short: **Tour de contrôle SI** |
| **🇨🇦 Canada FR** (`fr-CA`) — **42 Pages** | • Homepage: [servicenow.com/fr-ca/](https://www.servicenow.com/fr-ca/)<br>• Flagship: [/fr-ca/products/ai-control-tower.html](https://www.servicenow.com/fr-ca/products/ai-control-tower.html) | **Tour de contrôle Super Intelligence**<br>Short: **Tour de contrôle SI** |
| **🇪🇸 Spain** (`es-ES`) — **43 Pages** | • Homepage: [servicenow.com/es/](https://www.servicenow.com/es/)<br>• Flagship: [/es/products/ai-control-tower.html](https://www.servicenow.com/es/products/ai-control-tower.html) | **Torre de control de Superinteligencia**<br>Short: **Torre de control de SI** |
| **🌎 LATAM** (`es-419`) — **41 Pages** | • Homepage: [servicenow.com/latam/](https://www.servicenow.com/latam/)<br>• Flagship: [/latam/products/ai-control-tower.html](https://www.servicenow.com/latam/products/ai-control-tower.html) | **Torre de control de Superinteligencia**<br>Short: **Torre de control de SI** |
| **🇧🇷 Brazil** (`pt-BR`) — **44 Pages** | • Homepage: [servicenow.com/br/](https://www.servicenow.com/br/)<br>• Flagship: [/br/products/ai-control-tower.html](https://www.servicenow.com/br/products/ai-control-tower.html) | **Torre de controle de Superinteligência**<br>Short: **Torre de controle de SI** |
| **🇮🇹 Italy** (`it-IT`) — **41 Pages** | • Homepage: [servicenow.com/it/](https://www.servicenow.com/it/)<br>• Flagship: [/it/products/ai-control-tower.html](https://www.servicenow.com/it/products/ai-control-tower.html) | **Torre di controllo della Super Intelligenza**<br>Short: **Torre di controllo SI** |
| **🇳🇱 Netherlands** (`nl-NL`) — **45 Pages** | • Homepage: [servicenow.com/nl/](https://www.servicenow.com/nl/)<br>• Flagship: [/nl/products/ai-control-tower.html](https://www.servicenow.com/nl/products/ai-control-tower.html) | **Superintelligentie-controletoren**<br>Short: **SI-controletoren** |
| **🇯🇵 Japan** (`ja-JP`) — **46 Pages** | • Homepage: [servicenow.com/jp/](https://www.servicenow.com/jp/)<br>• Flagship: [/jp/products/ai-control-tower.html](https://www.servicenow.com/jp/products/ai-control-tower.html) | **スーパーインテリジェンス・コントロールタワー**<br>Short: **SI コントロールタワー** |
| **🇰🇷 Korea** (`ko-KR`) — **43 Pages** | • Homepage: [servicenow.com/kr/](https://www.servicenow.com/kr/)<br>• Flagship: [/kr/products/ai-control-tower.html](https://www.servicenow.com/kr/products/ai-control-tower.html) | **슈퍼 인텔리전스 컨트롤 타워**<br>Short: **SI 컨트롤 타워** |

---

## 5. Wave Rollout Summary (Tab 5 — No Hardcoded Dates)

Regional pages automatically inherit the Wave assignment of their US (`en-US`) parent page:

| Wave | Pages & Regional Split | Execution Split & Traffic | Included Page Groups |
| :--- | :--- | :--- | :--- |
| **Wave 1** | **414 Total Pages**<br>• 🇺🇸 54 US Pages<br>• 🌍 360 Regional Pages<br>*(1,904 total edits)* | • **🤖 282 Pure Auto / 🛠️ 132 Hybrid Manual**<br>• **5,500,465 US Visits** *(96.2%)*<br>• **38,805 Conversions** *(90.8%)* | Global Nav/Search XFs + US & all 13 Regional Homepages ([servicenow.com/](https://www.servicenow.com/), [servicenow.com/uk/](https://www.servicenow.com/uk/), [/de/](https://www.servicenow.com/de/), [/fr/](https://www.servicenow.com/fr/), etc.) + Flagship Product Pages ([/products/ai-control-tower.html](https://www.servicenow.com/products/ai-control-tower.html) & all 13 locales) + Top Traffic Hub Pages ([/ai.html](https://www.servicenow.com/ai.html), [/now-platform.html](https://www.servicenow.com/now-platform.html), etc.). |
| **Wave 2** | **353 Total Pages**<br>• 🇺🇸 173 US Pages<br>• 🌍 180 Regional Pages<br>*(759 total edits)* | • **🤖 307 Pure Auto / 🛠️ 46 Hybrid Manual**<br>• **202,226 US Visits** *(3.6%)*<br>• **3,777 Conversions** *(8.9%)* | All remaining Commercial Product, Solution, Capability, Application, Platform, Industry, Event, and Resource Center pages across US + 13 regional locales. |
| **Wave 3** | **166 Total Pages**<br>• 🇺🇸 136 US Pages<br>• 🌍 30 Regional Pages<br>*(347 total edits)* | • **🤖 161 Pure Auto / 🛠️ 5 Hybrid Manual**<br>• **12,822 US Visits** *(0.2%)*<br>• **130 Conversions** *(0.3%)* | Long-tail Archival Blogs, Low-traffic Workflow articles (<100 visits), Customer Stories, Partners, University, and Docs across US + 13 regional locales. |
| **Total** | **933 Total Pages**<br>• **🇺🇸 363 US**<br>• **🌍 570 Regional** | • **🤖 750 Pure Auto (94.9%)**<br>• **🛠️ 183 Hybrid Manual (5.1%)**<br>• **5,715,513 US Visits (100%)** | Complete global + regional rebrand handoff inventory. |
