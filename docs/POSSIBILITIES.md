# My Work Hub — Possibilities & scope

Decisions locked: **name = My Work Hub**, **recent work = item insights**, **Copilot default = “Summarize my pending work today”**, **tab order = Tasks first**. This doc records scope and what to build first vs later.

---

## 1. Tasks / actions hub

### Must-have (v1)

- **To Do** — List task lists; list tasks per list; show due date, title, completed; link to open in To Do/Outlook/Teams.
- **Planner** — List tasks assigned to me; show plan name, bucket, due date, title; link to open in Planner/Teams.
- **Unified list** — Either one merged list (sorted by due date or priority) or two groups “To Do” and “Planner”.
- **Mark complete** — Where Graph supports it (To Do: PATCH task; Planner: PATCH task).

### Nice-to-have

- **Flagged emails** — If Graph exposes “flagged messages” in a way we can list (e.g. filter on flag), add a third group “Flagged” with link to open in Outlook. Otherwise: a single “Open Outlook flagged items” link.
- **Filters** — By due date (overdue, today, this week), by source (To Do / Planner).
- **Create task** — Quick “Add task” that creates in a default To Do list (optional for hackathon).

### Out of scope for v1

- Outlook Tasks (deprecated); other task systems.

---

## 2. My approvals hub

### Must-have (v1)

- **List pending approvals** — Approvals app (Teams/Power Automate) via Graph beta: list `approvalItem` objects; filter to “assigned to me” and state = pending (client-side filter if needed).
- **Display** — Title, description (truncated), requestor, created date.
- **Approve / Reject** — Call Graph to create `approvalItemResponse`; refresh list after action.
- **Link** — Open in Teams Approvals app if we have a deep link.

### Nice-to-have

- **Custom responses** — For custom approval types, show custom response options and send chosen response.
- **Group by requestor or date** — Optional view.
- **Badge count** — e.g. “3 pending” in tab or header.

### Out of scope for v1

- Entitlement Management / PIM approvals (different API); creating new approval requests from the web part.

---

## 3. My recent work

### Must-have (v1)

- **Item insights** — Use Graph **item insights** API for recommended/relevant files (decided; avoids deprecated `me/drive/recent`).
- **Display** — File name, site/location (OneDrive vs SharePoint), last modified; link to open in browser.
- **Limit** — Top 10–20 items.

### Nice-to-have

- **Sub-sections** — “Recommended for you” vs “Recent” if the insights API supports both.
- **Thumbnails** — If Graph returns thumbnails and we can show them without much overhead.
- **Activity** — `me/activities/recent` for a broader “recent activity” (e.g. edited pages, comments). Optional.

### Out of scope for v1

- Full activity feed; editing files from the web part.

---

## 4. Copilot spice

### Must-have (v1) — decided

- **Default prompt:** **“Summarize my pending work today.”**  
  Summary tab: on load (or “Summarize” button) we build a short context string from Tasks, Approvals, and Recent (e.g. “Tasks: 3 overdue, 5 due this week. Pending approvals: 2. Recent files: 5.”), send to **Chat API** with this prompt, and show the reply in a card or panel.

### Nice-to-have

- **Follow-up questions** — User can ask “What’s the most urgent?” or “Summarize the first approval” in the same Chat session.
- **Summarize approval** — For selected approval, send title + description to Chat: “Summarize this approval request and suggest approve or reject with one reason.”
- **Smart grouping** — “Group my tasks by project” (we’d need to infer “project” from plan name or list name; or ask Chat to suggest labels).

### Out of scope for v1

- Retrieval over full document content for “recent work” (heavy); custom Copilot plugins.

---

## 5. UX and layout

### Must-have (v1)

- **Tabs (order decided):** **Tasks | Approvals | Recent | Summary.** Summary = Copilot “Summarize my pending work today.”
- **Loading and errors** — Spinners per section; graceful message if a Graph or Copilot call fails (e.g. “Approvals couldn’t load”).
- **Responsive** — Usable on desktop and tablet (single column on small screens).
- **Theme** — Use Fluent UI / SharePoint theme so it fits the host page.

### Nice-to-have

- **Configurable visibility** — Property pane: “Show Tasks: yes/no”, “Show Approvals: yes/no”, etc., so site owners can hide sections.
- **Unified feed** — Optional view: one list with tasks, approvals, and recent files mixed, sorted by date.
- **Badges** — Counts on tabs (e.g. “Approvals (3)”).

---

## 6. Technical constraints (summary)

- **Approvals:** Graph **beta** only (Approvals app); list and respond to approvalItem.
- **Recent work:** We use **item insights** (decided); `me/drive/recent` is deprecated Nov 2026.
- **Copilot:** **Chat API** is preview; requires M365 Copilot license; we use REST via Graph (`/copilot/...`).
- **SPFx:** Single web part; MSGraphClient or custom Graph client with delegated permissions.

Details and permission list: [TECHNICAL_NOTES.md](./TECHNICAL_NOTES.md).

---

## 7. MVP (for hackathon) — locked

| Component | MVP choice |
|-----------|------------|
| **Tasks** | To Do + Planner, merged list, mark complete, link to open. **Tab 1.** |
| **Approvals** | List pending (Approvals app), approve/reject, link to Teams. **Tab 2.** |
| **Recent work** | **Item insights** (top 10–15). **Tab 3.** |
| **Copilot** | **“Summarize my pending work today”** via Chat API with aggregated context. **Tab 4: Summary.** |
| **Layout** | 4 tabs in order: **Tasks | Approvals | Recent | Summary.** |
| **Polish** | Loading states, error handling, Fluent UI, basic responsive. |

Single web part **My Work Hub** combining the three pillars and Copilot. Future: flagged emails, custom approval responses, more Copilot prompts.

---

## 8. Decisions (resolved)

- [x] **Recent work source:** **Item insights.** (Avoids deprecated `me/drive/recent`; better relevance.)
- [x] **Copilot default prompt:** **“Summarize my pending work today.”**
- [x] **Tab order:** **Tasks first** → Approvals → Recent → Summary.
- [x] **Name:** **My Work Hub.**
