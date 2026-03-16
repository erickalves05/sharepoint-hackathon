# SharePoint Hackathon 2026 — New Submissions Analysis

**Analysis date:** March 2026  
**Scope:** Submissions **#143–#176** (new since the previous analysis, which covered #118–#142).  
**Source:** [SharePoint Hackathon Issues](https://github.com/SharePoint/sharepoint-hackathon/issues)

---

## Summary

| Metric | Count |
|--------|--------|
| New submissions in range | **34** (issues #143–#176; some issue numbers may be missing) |
| With **SharePoint Framework** label | 12+ |
| With **Agents and SharePoint** | 14+ |
| With **Knowledge Agent** | 10+ |
| With **SharePoint Embedded** | 3 |
| With **SharePoint site** (design/OOB) | 14+ |
| With **Mobile experience** | 2 |

**Notable trends:** Strong cluster around **Knowledge Agent** and **Agents + SharePoint**; several **SPFx + Copilot/AI** solutions; a few **SharePoint Embedded**; design/OOB intranets and vertical solutions (claims, policy, onboarding, sales).

---

## New submissions by category

### SharePoint Framework (SPFx)

| # | Project | One-line idea | Labels |
|---|--------|----------------|--------|
| **176** | **Template Deployer** | Deploy page templates to many sites in one click; Graph for site discovery; conflict detection, retry, deployment log (CSV). SPFx 1.22.2, no external infra. | SPFx |
| **175** | **TransmittalFlow** | SPFx Command Set + SharePoint Embedded: select docs → transmittal → isolated container per client; external portal (no M365) for AI summary, co-author, approve/reject; status back to SharePoint. | SPFx, Embedded |
| **174** | **Day One** | Intelligent onboarding: SPFx Journey Tracker (weeks 1–8, role/department filter via Graph, completion to lists, confetti); Copilot Studio agent “Ask HR” grounded on site; Sprocket 365. | Site, Agents, SPFx |
| **169** | **SharePoint OS** | “OS-style” digital workspace in one click: File Explorer, Recent Files, SharePoint Assistant, Lists, Notes, To-Do, etc. SPFx + Agents. | Agents, SPFx |
| **167** | **Policy Consistency Checker** | SPFx web part: on doc upload, chunk → embed → Cosmos vector search → LLM judge for contradictions; dashboard, side-by-side, Copilot Chat (ask, Q&A, reconciliation, weekly digest). Azure Functions, Redis, webhooks. | Site, Agents, SPFx |
| **164** | **K-Docs Publish** | SPFx: publish Word docs as Site Pages for knowledge base; tree nav, style mappings; RAG search with Azure AI Foundry. | Knowledge Agent, Agents, SPFx |
| **145** | **Copilot Engagement Program** | SPFx web parts + App Customizer: opt-in dashboard, Copilot Wins, leaderboard; Teams personal app + declarative agent to capture wins; Azure Functions, Copilot usage APIs, badges. | Agents, SPFx |
| **144** | **ShareGPT** | SPFx RAG chatbot: Power Automate → Blob → FastAPI (chunk, embed, Azure AI Search) → SPFx chat; citations, guardrails, Cosmos for history; SharePoint + Teams. | Site, Knowledge Agent, Agents, SPFx |

### Knowledge Agent (library / content intelligence)

| # | Project | One-line idea | Labels |
|---|--------|----------------|--------|
| **173** | **Intelligent Knowledge Hub (Courier)** | Knowledge Agent: auto-analyze uploads, extract metadata, organize; summaries, stakeholders, rules for organizing; courier operations scenario. | Knowledge Agent |
| **172** | **AI Sales Playbook** | Knowledge Agent on sales library: auto metadata (industry, stage, etc.), natural language rules, context-aware Q&A; view formatting for expiring policies, case studies. | Knowledge Agent |
| **168** | **AI-Powered Policy Intelligence Center** | Policy library + metadata; Policy Intelligence Agent (questions on expiry, impact, owner); Power Automate metrics → list → dashboard. OOB + agent. | Site, Knowledge Agent, Agents |
| **165** | **Aura** | AI-ready knowledge hub: Ask / Understand / Resolve / Act; structured articles, metadata (topic, owner, review, AI summary, confidence); AI autofill, Copilot retrieval. OOB + AI in SharePoint. | Site, Knowledge Agent, Agents |
| **164** | **K-Docs Publish** | (see SPFx) Word → Site Pages, RAG. | Knowledge Agent, Agents, SPFx |
| **144** | **ShareGPT** | (see SPFx) Enterprise RAG chatbot. | Site, Knowledge Agent, Agents, SPFx |

### Agents and SharePoint (Copilot / agents)

| # | Project | One-line idea |
|---|--------|----------------|
| **174** | **Day One** | Copilot Studio “Ask HR” + SPFx Journey Tracker. |
| **170** | **CAP Claim Xpress** | Power Apps + Copilot (Xpress Agent) + Syntex; claims, KYC, policy Q&A; broker portal; optional Embedded. |
| **169** | **SharePoint OS** | “OS” workspace with Assistant, Lists, To-Do, etc. |
| **168** | **Policy Intelligence Center** | Agent on policy library + automation + dashboard. |
| **167** | **Policy Consistency Checker** | Copilot Chat for conflict analysis, Q&A, reconciliation, digest. |
| **165** | **Aura** | Guided knowledge + AI autofill + Copilot. |
| **164** | **K-Docs Publish** | RAG with Azure AI Foundry. |
| **150** | **SPARK – AI Solution Accelerator** | Copilot Studio: PRD → schema → Power Automate (Graph) provisioning → validation engine → Mermaid diagrams. No SPFx. |
| **145** | **Copilot Engagement Program** | SPFx + declarative agent + Copilot usage APIs. |
| **144** | **ShareGPT** | RAG + SPFx chat. |
| **143** | **MilkNet AI Workplace** | Intranet + AI news, productivity hub, persona-based agents. |

### SharePoint Embedded

| # | Project | One-line idea |
|---|--------|----------------|
| **175** | **TransmittalFlow** | External client portal (no M365) over Embedded containers; co-author, approve/reject. |
| **170** | **CAP Claim Xpress** | Listed Embedded (portal/experience). |

### Design / SharePoint site (OOB, intranet, mobile)

| # | Project | One-line idea | Labels |
|---|--------|----------------|--------|
| **171** | **Batcave (An Intranet Experience)** | Themed intranet (Gotham/Batman) with OOB web parts only: header, news, announcements, Excel in page, editorial cards, quick links, countdown, people, map. | Site, Mobile |
| **174** | **Day One** | Onboarding site + Sprocket 365 + SPFx + agent. | Site, Agents, SPFx |
| **165** | **Aura** | Knowledge hub design + metadata + AI. | Site, Knowledge Agent, Agents |
| **168** | **Policy Intelligence Center** | Policy site + agent + dashboard. | Site, Knowledge Agent, Agents |
| **143** | **MilkNet AI Workplace** | Dairy intranet, AI news, productivity hub, agents. | Site, Knowledge Agent, Agents |

### Other (vertical / low-code)

| # | Project | One-line idea |
|---|--------|----------------|
| **170** | **CAP Claim Xpress** | Insurance claims: Power Apps, Copilot, Syntex, KYC, broker workflow; optional Embedded. |
| **150** | **SPARK** | PRD → Copilot Studio → Graph provisioning → validation; no SPFx. |

---

## Thematic clusters

### 1. **Productivity / “one place” / OS-style**

- **SharePoint OS (#169)** — File Explorer, Recent Files, Assistant, Lists, Notes, To-Do in one workspace.  
- **My Work Hub (our concept)** — Tasks (To Do + Planner), Approvals, Recent work + Copilot fits here; distinct by focusing on **tasks + approvals + recent** with Graph + Approvals API + one Copilot “focus” experience.

**Differentiation:** We’re not building a full “OS” shell; we’re a single web part that aggregates **actions + approvals + recent** with a clear Copilot layer. No direct overlap with SharePoint OS’s broader module set.

### 2. **Policy / governance / compliance**

- **Policy Consistency Checker (#167)** — Contradiction detection (vector + LLM) + Copilot resolution.  
- **AI-Powered Policy Intelligence Center (#168)** — Agent on policy library + metrics + dashboard.  
- **Aura (#165)** — Knowledge hub pattern applicable to policy/onboarding.

**Gap:** No submission is a **personal productivity hub** (tasks + approvals + recent) like My Work Hub.

### 3. **Knowledge Agent / RAG**

- **ShareGPT (#144)**, **K-Docs Publish (#164)**, **Aura (#165)**, **Knowledge Hub (#173)**, **Sales Playbook (#172)** — Library-centric, metadata, Q&A, RAG.  
- My Work Hub is **user-centric** (my tasks, my approvals, my recent work), not document-centric.

### 4. **Onboarding**

- **Day One (#174)** — Onboarding journey (weeks 1–8), Ask HR agent, Sprocket 365.  
- Different use case from My Work Hub (daily productivity, not first-week onboarding).

### 5. **Deployment / provisioning**

- **Template Deployer (#176)** — One-click template deploy across sites.  
- **SPARK (#150)** — PRD → schema → Graph provisioning.  
- No overlap with My Work Hub.

### 6. **External / Embedded**

- **TransmittalFlow (#175)** — Transmittals + external portal (Embedded).  
- **CAP Claim Xpress (#170)** — Claims/KYC with optional Embedded.  
- My Work Hub is internal (user’s own tasks/approvals/recent).

### 7. **Engagement / adoption**

- **Copilot Engagement Program (#145)** — Opt-in Copilot usage, wins, leaderboard, SPFx + agent.  
- My Work Hub is about **getting work done** (tasks, approvals, recent), not gamifying Copilot usage.

---

## Implications for “My Work Hub”

| Question | Conclusion |
|----------|------------|
| **Is there another “tasks + approvals + recent” hub?** | **No.** SharePoint OS (#169) has Lists/To-Do/Recent in a broader “OS” shell but not a dedicated **approvals hub** or a single web part focused on these three pillars. |
| **Is Copilot overused?** | Many use Copilot/agents for **knowledge** (Q&A, RAG, policy). Few use it for **personal prioritization** (“What should I focus on?” over tasks + approvals + recent). Our angle remains distinct. |
| **SPFx + Graph + Copilot still a strong combo?** | **Yes.** Policy Consistency Checker (#167), ShareGPT (#144), Copilot Engagement (#145), Day One (#174), K-Docs (#164) all use SPFx with Graph and/or Copilot. |
| **Design / OOB competition?** | Batcave (#171), Aura (#165), Day One (#174) show strong design/OOB. My Work Hub competes mainly on **SPFx** and **Agents** (Copilot); design is secondary but should be clean (Fluent, tabs, responsive). |
| **Risk of “yet another dashboard”?** | Reduced if we **clearly own** (1) **Approvals in one place** (Graph Approvals app) and (2) **Copilot “focus” over tasks+approvals+recent**. No other submission does both in one web part. |

---

## Recommended positioning for My Work Hub

1. **Single web part:** “Tasks (To Do + Planner) + My approvals (Teams Approvals app) + Recent work” — **one place to start the day.**  
2. **Copilot:** “What should I focus on?” or “Summarize my pending work” using **Chat API** and context built from the three sections.  
3. **Technical hook:** Use of **Graph Approvals app API** (beta) for list + approve/reject in SharePoint is still rare among submissions.  
4. **Category:** Submit under **SharePoint Framework** and **Agents and SharePoint**; stress **productivity** and **Copilot for prioritization**, not document knowledge.

---

## Quick reference: all new submission issue numbers

| # | Title (short) |
|---|----------------|
| 176 | Template Deployer |
| 175 | TransmittalFlow |
| 174 | Day One – Intelligent Employee Onboarding |
| 173 | Intelligent Knowledge Hub (Courier) |
| 172 | AI Powered Sales Playbook |
| 171 | Batcave – An Intranet Experience |
| 170 | CAP Claim Xpress |
| 169 | SharePoint OS |
| 168 | AI-Powered Policy Intelligence Center |
| 167 | Policy Consistency Checker |
| 165 | Aura – AI-Ready Knowledge Hub |
| 164 | K-Docs Publish |
| 150 | SPARK – AI SharePoint Solution Accelerator |
| 149 | FollowUp Plus |
| 148 | DecisionForge |
| 147 | ResponseIQ |
| 146 | Corporate Gateway |
| 145 | Copilot Engagement Program |
| 144 | ShareGPT |
| 143 | MilkNet AI Workplace |

*(Issues 151–163 not listed on the first two pages of Issues; may be closed or different filters.)*

---

**Next step:** Proceed with My Work Hub as planned; no pivot needed. Emphasize **approvals hub + Copilot “focus”** in the submission narrative to stand out.
