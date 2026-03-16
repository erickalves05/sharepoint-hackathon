# My Work Hub — Concept

## Vision

**One web part = one place** for a user’s immediate work: tasks, approvals, and recent work, with Copilot helping prioritize and summarize.

Today, this is spread across:
- To Do, Planner, Outlook (tasks / flags)
- Teams Approvals, Power Automate (approvals)
- OneDrive, SharePoint, Outlook (recent files and activity)

**My Work Hub** aggregates these into a single SPFx web part on a SharePoint page (or Teams) so users start their day from one view. The default Copilot prompt is **“Summarize my pending work today.”**

---

## User value

1. **Less context switching** — Tasks, approvals, and recent work in one place.
2. **Clear priorities** — Overdue tasks, pending approvals, and “trending” recent work visible at a glance.
3. **Faster actions** — Approve/reject, complete tasks, open recent files without leaving the page.
4. **AI assist** — Copilot suggests focus, summarizes the list, or answers questions about “my” work (grounded in the same data).

---

## The three pillars (in one web part)

### 1. Tasks / actions hub

- **To Do** — Task lists and tasks (Graph: `me/todo/lists`, `me/todo/lists/{id}/tasks`).
- **Planner** — Tasks assigned to me (Graph: `me/planner/tasks`).
- **Flagged emails (optional)** — If available via Graph (e.g. Outlook mail with flag); otherwise defer or show “Open Outlook” for flagged.

**UX:** Single list or grouped by source (To Do / Planner / Flagged). Due date, title, source, link to open. Mark complete where the API allows (To Do, Planner).

### 2. My approvals hub

- **Approvals app (Teams / Power Automate)** — Approval items where I am an approver (Graph beta: Approvals app API, `approvalItem`).
- Show: title, description, requestor, date, state (pending/completed).
- Actions: Approve / Reject (and custom responses if supported) from the web part via Graph.

**UX:** List of pending approvals; expand or open detail for description and response options; one-click approve/reject where possible.

### 3. My recent work

- **Item insights** — “Recommended” / relevant files for the user (Graph: item insights API). Chosen over `me/drive/recent` (deprecated Nov 2026) for better long-term relevance.
- **Optional:** `me/activities/recent` later for a broader “what I touched” feed.

**UX:** Cards or list: file name, site/location, last modified, thumbnail if easy. Click to open in new tab.

---

## Copilot spice

**MVP (v1):** One Copilot feature using the **Chat API** with a small context blob (task titles, approval titles, recent file names from the other tabs).

**Default prompt (decided):** **“Summarize my pending work today.”** — On load of the Summary tab (or on demand), we send aggregated context to the Chat API with this prompt and show the reply in a card or panel.

**Future options:** “What should I focus on?”, smart grouping, natural-language search, summarize a selected approval.

---

## One web part, layout

- **Tabs (v1, order decided):** **Tasks | Approvals | Recent | Summary.**  
  Summary tab runs the default Copilot prompt: “Summarize my pending work today.”
- **Optional later:** Single feed (mixed stream), or configurable tab order.

---

## Scope summary

| Include in v1 | Defer or optional |
|---------------|--------------------|
| To Do + Planner tasks in one list | Flagged emails (or link to Outlook) |
| Approvals app: list + approve/reject | Other approval types (e.g. Entitlement Management) |
| Recent work via **item insights** | Full activity feed |
| Copilot: **“Summarize my pending work today”** (Summary tab) | Multiple Copilot scenarios |
| Tabs (order): **Tasks | Approvals | Recent | Summary** | Single unified feed |

---

## Name

**My Work Hub** — “My place to start the day” / “Everything in one place.” Short and memorable for the hackathon submission.
