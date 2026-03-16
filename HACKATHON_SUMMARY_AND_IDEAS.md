# SharePoint Hackathon 2026 – Summary & Project Ideas

## Hackathon overview (from README)

- **When:** March 2–16, 2026 | **Submissions due:** Monday March 16, 2026, 11:59 PM PST  
- **Open to:** Everyone (end users, designers, architects, developers)  
- **Submit:** Video (max 8 min) + GitHub Issue via [Project Submission Form](https://aka.ms/SharePoint/Hackathon/ProjectSubmission)  
- **Awards:** March 25, 2026  

### Categories relevant to you

| Category | What it rewards |
|----------|------------------|
| **Best design for SharePoint Site** | Outstanding visual design, branding, layout, usability – polished, engaging experience. |
| **Most innovative SharePoint experience with SPFx** | SPFx solutions that extend SharePoint UI: custom web parts, extensions, experiences beyond OOB. *Extra points for AI (e.g. Copilot APIs).* |

One submission can enter multiple categories. Judges consider innovation, impact, technical usability, and fit to the category.

---

## Current submissions (from Issues)

### SPFx-focused (SharePoint Framework)

| # | Project | One-line idea | Tech / angle |
|---|--------|----------------|--------------|
| **136** | **Search Booster** | AI search that understands multiple languages and summarizes results | Copilot Retrieval + Conversations API, multilingual, SPFx + React + Fluent UI |
| **135** | **LegalLens** | Contract management + e-signature in SharePoint | SPFx, Azure AI, Document Intelligence, e-sign, risk scoring, Q&A agent |
| **134** | **JAKMAN ISO Management System** | ISO compliance (Quality, Env, OHS) as SPFx app | Application Customizer, List Command Set, web parts, list formatting, “content installation” |
| **133** | **SPFx Site Game** | Turn site content into a 2D discovery game (lists/libraries = buildings) | Gamification, themes, discoverability |
| **132** | **Bookmark Hub** | One place for flagged emails, favorites, followed sites, custom URLs | Graph API, groups/labels, Copilot to suggest groups |
| **131** | **AI requirement gathering & auto-provisioning** | Describe app in chat → AI gathers requirements → auto-creates lists, groups, pages | SPFx chat bot + Azure AI Agent + dynamic template web part |
| **130** | **Mermaid Diagram from page content** | Turn page text into Mermaid diagrams via Copilot API | M365 Copilot API, one-click + refinement feedback, license check |
| **128** | **Company Directory – Powered by AI** | (SPFx) | AI-powered directory experience |
| **124** | **Workstream Oasis** | (SPFx + Agents + Site) | Agents + SPFx + site |
| **119** | **Smart Export to PDF** | SPFx export to PDF in SharePoint Online | SPFx, PDF generation |
| **118** | **Back to the Future with Microsoft Copilot** | (SPFx + Site) | Copilot + SPFx |

### Design / site-focused (no or light SPFx)

| # | Project | One-line idea |
|---|--------|----------------|
| **138** | **One-Click User Friendly Intranet (Art of Adoption)** | Intranet with OOB web parts only: user-first, easy to maintain, action-ready; Brand Center, audience targeting, highlighted content |
| **142** | **Automates entire intranet setup** | Single-button intranet setup |
| **141** | **Offline SP Form Customizer** | Form customizer that works offline |
| **140** | **Power Platform User Group SharePoint Knowledge Hub** | Knowledge hub on SharePoint |
| **137** | **Smart Campus 365** | AI-powered school intranet |

### Other (SPFx + design angle)

| # | Project | Note |
|---|--------|------|
| **139** | **SymanticRTE** | Semantic rich text editor SPFx (typography, design system, Brand Center, no React – HTML/CSS/JS). Strong **design + SPFx** story. |

---

## Patterns and gaps (ideas you could build)

### 1. **SPFx + Copilot / AI (judges like this)**

- **Already in:** Search (Retrieval + Chat), Legal Q&A, requirement gathering, Mermaid from text, Bookmark Hub (suggest groups), Company Directory.  
- **Gaps you could fill:**
  - **Page / content summarizer:** “Summarize this page” or “Summarize this section” button with Copilot API.
  - **Smart templates:** “Generate a meeting recap / project brief / policy summary from this doc” as SPFx action.
  - **Accessibility / readability:** SPFx web part that uses AI to suggest simpler language or alternative formats.
  - **Search + visualization:** SPFx that uses Copilot retrieval then charts/timelines (e.g. “show me projects by date”).
  - **Localization / terminology:** SPFx that suggests consistent terms or translations for page/list content using Copilot.

### 2. **Productivity & “one place” (like Bookmark Hub)**

- **Already in:** Bookmark Hub (emails, files, sites, URLs in one place).  
- **Gaps:**
  - **Tasks / actions hub:** One web part for Tasks, Planner, To Do, flagged emails, with filters and “focus” view.
  - **“My approvals” hub:** All approval types (Power Automate, Lists, etc.) in one SPFx view.
  - **“My recent work” dashboard:** Recent files, pages, sites, and comments in one customizable SPFx dashboard.

### 3. **Content creation & authoring (like SymanticRTE, Mermaid)**

- **Already in:** Semantic RTE, Mermaid from text.  
- **Gaps:**
  - **Structured content from free text:** “Turn this bullet list into a checklist” or “Turn this paragraph into a table” (Copilot API).
  - **Image alt-text and captions:** SPFx that suggests or generates alt-text and captions for images on the page.
  - **Glossary / definitions:** Web part that detects terms and links them to a SharePoint glossary list or generates tooltips.

### 4. **Design-first (Best design for SharePoint Site)**

- **Already in:** Art of Adoption (OOB only), SymanticRTE (typography, Brand Center).  
- **Gaps:**
  - **Design system showcase site:** A single SharePoint site that demonstrates one clear design system (color, type, spacing, components) using OOB + optional small SPFx “wrapper” web parts.
  - **Accessible theme builder:** SPFx or OOB-based way to preview and export contrast-safe themes (Brand Center + APCA, like SymanticRTE).
  - **Template gallery:** Beautiful, ready-to-use page templates (OOB + optional SPFx) for common scenarios (team home, project, policy, event) with strong visual design and clear layout rules.

### 5. **Domain-specific SPFx (like LegalLens, JAKMAN)**

- **Already in:** Legal (contracts, e-sign), ISO/compliance, Campus.  
- **Gaps:**
  - **HR onboarding / offboarding checklist** as SPFx (tasks, links, docs, approvals in one place).
  - **Project health dashboard** (risks, milestones, deliverables from lists + optional AI summary).
  - **Vendor / partner portal** (document exchange, simple e-sign, status) without going as broad as LegalLens.

### 6. **Gamification & engagement (like SPFx Site Game)**

- **Already in:** 2D site discovery game.  
- **Gaps:**
  - **Training / adoption tracker:** SPFx that shows progress (e.g. “visited 5 key sites”, “completed 3 trainings”) with simple badges or progress bar.
  - **“Site tour” or “feature spotlight”** as a small SPFx overlay or guided path.

### 7. **Forms & data (like Offline Form Customizer)**

- **Already in:** Offline form customizer.  
- **Gaps:**
  - **Form customizer** that adds logic (show/hide, validation, default values from Graph) with a clear UX.
  - **List form that writes to multiple lists** or syncs to another system, with a clean SPFx form UI.

---

## Suggested directions for you

**If you prefer SPFx:**

1. **SPFx + Copilot in a narrow scenario**  
   Pick one clear use case (e.g. “summarize this page”, “suggest tags”, “turn list into chart”) and do it really well.  
2. **One “hub” web part**  
   Like Bookmark Hub but for a different set of things (approvals, tasks, or “my recent work”).  
3. **Authoring aid**  
   E.g. Mermaid from text, alt-text generator, or “bullet list → checklist” – small scope, clear value.

**If you prefer design:**

1. **One outstanding SharePoint site**  
   Use only OOB features + Brand Center (like Art of Adoption) but with a very strong visual theme and clear design principles.  
2. **Design system + one SPFx**  
   Define a small design system (type scale, colors, spacing) and build one SPFx web part that showcases it (e.g. semantic editor, card layout, or dashboard).  
3. **Template pack**  
   A set of 3–5 page templates with clear layout and styling, documented so others can replicate.

**If you want both SPFx and design:**

- **SymanticRTE-style project:** One SPFx that is both useful (e.g. better authoring, readability, or structure) and visually/typographically strong, with Brand Center integration and accessibility in mind.

---

## Quick links

- **README:** https://github.com/SharePoint/sharepoint-hackathon  
- **Official rules:** https://github.com/SharePoint/sharepoint-hackathon/blob/main/OFFICIAL_RULES.md  
- **Project submission:** https://aka.ms/SharePoint/Hackathon/ProjectSubmission  
- **Badge form:** https://aka.ms/SharePoint/Hackathon/Badges  
- **Discussions:** https://github.com/SharePoint/sharepoint-hackathon/discussions  
- **SPFx:** https://aka.ms/spfx  
- **Copilot APIs:** https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/copilot-apis-overview  

Good luck with the hackathon.
