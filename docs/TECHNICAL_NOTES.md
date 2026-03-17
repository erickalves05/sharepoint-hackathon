# My Work Hub — Technical notes

**Product name:** My Work Hub. **Decisions:** Recent work = item insights; Copilot default = “Summarize my pending work today”; tab order = Tasks | Approvals | Recent | Summary.

References and implementation notes from **Microsoft Learn** and Graph/Copilot docs. Use this when scaffolding the SPFx app and calling APIs.

---

## 1. SharePoint Framework (SPFx)

- **Docs:** [Develop web parts with the SharePoint Framework](https://learn.microsoft.com/en-us/training/modules/sharepoint-spfx-web-parts/), [Build your first web part](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part).
- **Stack:** SPFx 1.x, TypeScript, React, Fluent UI. Use `MSGraphClient` from `@microsoft/sp-http` for Graph calls in SharePoint context (delegated permissions).
- **Testing:** Hosted workbench `https://{tenant}.sharepoint.com/sites/{site}/_layouts/15/workbench.aspx`, `gulp serve`.
- **Package:** `.sppkg` to App Catalog; deploy to site.

---

## 2. Microsoft Graph — Tasks

### To Do (primary for personal tasks)

- **Overview:** [To Do API overview](https://learn.microsoft.com/en-us/graph/todo-concept-overview).
- **Endpoints (v1.0):**
  - List task lists: `GET https://graph.microsoft.com/v1.0/me/todo/lists`
  - List tasks: `GET https://graph.microsoft.com/v1.0/me/todo/lists/{todoTaskListId}/tasks?$expand=linkedResources`
  - Update task (e.g. complete): `PATCH https://graph.microsoft.com/v1.0/me/todo/lists/{todoTaskListId}/tasks/{todoTaskId}` — set `status: completed`.
- **Permissions:** `Tasks.Read`, `Tasks.ReadWrite` (delegated).
- **Notes:** Outlook Tasks API is deprecated; use To Do only. **Flagged emails** appear as To Do tasks from the default list; use `linkedResources[0].webUrl` to open the mail. Strip title suffix `(ID: ...)` for display.

### Planner (group/team tasks)

- **Overview:** [Planner REST API](https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0), [Planner tasks and plans](https://learn.microsoft.com/en-us/graph/planner-concept-overview).
- **Endpoint:** Tasks assigned to me: `GET https://graph.microsoft.com/v1.0/me/planner/tasks`
- **Plan name:** `GET https://graph.microsoft.com/v1.0/planner/plans/{planId}?$select=title` to resolve plan title for display.
- **Complete task:** GET task for `@odata.etag`, then `PATCH` with `If-Match` header and `percentComplete: 100`. Do **not** use `$select=@odata.etag` (invalid).
- **Permissions:** `Tasks.Read`, `Tasks.ReadWrite` (delegated).
- **Response:** `plannerTask` (title, planId, dueDateTime, percentComplete, etc.).

---

## 3. Microsoft Graph — Approvals (Teams Approvals app)

- **Overview:** [Approvals app API](https://learn.microsoft.com/en-us/graph/approvals-app-api).
- **Resource:** [approvalItem](https://learn.microsoft.com/en-us/graph/api/resources/approvalitem?view=graph-rest-beta) (beta only).
- **Endpoints (beta):**
  - List approval items: `GET https://graph.microsoft.com/beta/solutions/approval/approvalItems?$orderby=createdDateTime desc&$top=100` — **$filter on state is not allowed**; filter pending client-side. See [approvalsolution-list-approvalitems](https://learn.microsoft.com/en-us/graph/api/approvalsolution-list-approvalitems?view=graph-rest-beta).
  - Get one: `GET https://graph.microsoft.com/beta/solutions/approval/approvalItems/{id}`.
  - Create response (approve/reject): `POST https://graph.microsoft.com/beta/solutions/approval/approvalItems/{id}/responses` — [approvalItem-post-responses](https://learn.microsoft.com/en-us/graph/api/approvalitem-post-responses?view=graph-rest-beta).
- **Permissions:** `ApprovalSolution.Read` (read), `ApprovalSolution.ReadWrite` or `ApprovalSolutionResponse.ReadWrite` (to respond). Delegated only; no personal accounts.
- **Filtering “assigned to me”:** Use `$filter=state eq 'pending'` to get pending items. The API returns items where the user is owner or approver; each item has `viewPoint.roles` (e.g. `["Owner"]` or approver). Filter client-side for items where the current user is in approvers (or viewPoint indicates approver role) if needed. US Government L4/L5 and China clouds not supported (global only).
- **Note:** Different from Entitlement Management / PIM approvals (`identityGovernance/entitlementManagement/...`). We only use **solutions/approval** for Teams/Power Automate approvals.

---

## 3.5. Microsoft Graph — Calendar (Meetings)

- **Endpoint:** `GET https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '{isoNow}'&$orderby=start/dateTime&$top=5&$select=subject,start,end,organizer,webLink`
- **Permissions:** `Calendars.Read` (delegated).
- **Response:** Event with `subject`, `start.dateTime`, `end.dateTime`, `organizer.emailAddress.name`, `webLink`. Display with date badges (Today/Tomorrow), time range, organizer. Click opens `webLink`.

---

## 4. Microsoft Graph — Recent work (My Work Hub: item insights)

**Implemented:** My Work Hub uses **`/me/drive/recent`** for the Recent tab.

### Item insights (our choice)

- **Overview:** [Item insights in Microsoft Graph](https://learn.microsoft.com/en-us/graph/item-insights-overview).
- **Use:** “Recommended files” and relevance-based files for the user. Powers “Recommended” in M365.com and profile cards.
- **API:** [Insights API (beta)](https://learn.microsoft.com/en-us/graph/api/resources/iteminsights?view=graph-rest-beta), [v1.0](https://learn.microsoft.com/en-us/graph/api/resources/iteminsights?view=graph-rest-1.0). Explore `trending`, `used`, `shared` and related endpoints.
- **Permissions:** Verify in the insights API reference (e.g. `Sites.Read.All` or per-endpoint docs).

### Recent files (implemented)

- **Endpoint:** `GET https://graph.microsoft.com/v1.0/me/drive/recent?$top=15` — Returns DriveItem array. Response has `name`, `webUrl`, `lastModifiedDateTime`, `createdBy`, `file.mimeType`, `remoteItem` for shared. Map mimeType to icons; show creator, last modified.

### User activity (optional)

- **Endpoint:** `GET https://graph.microsoft.com/v1.0/me/activities/recent` (or beta).
- **Docs:** [Get recent user activities](https://learn.microsoft.com/en-us/graph/api/projectrome-get-recent-activities?view=graph-rest-beta).
- **Permissions:** `UserActivity.ReadWrite.CreatedByApp` (or as per docs). Broader “what I touched” feed.

---

## 5. Microsoft 365 Copilot APIs

- **Overview:** [Microsoft 365 Copilot APIs overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/copilot-apis-overview).
- **Integration:** REST under Microsoft Graph (`graph.microsoft.com/v1.0/copilot` or `.../beta/copilot`). Same auth as Graph (delegated with Copilot license).
- **Relevant for My Work Hub:**
  - **Chat API (preview)** — Conversational experience grounded in M365 data. Use for “What should I focus on?” or “Summarize my pending work” by sending a prompt and optional context (e.g. our aggregated task/approval/recent summary).
  - **Retrieval API (GA)** — Semantic search over SharePoint/OneDrive; optional if we want to ground answers in documents. For v1, Chat with a small context string is enough.
- **Licensing:** User needs **Microsoft 365 Copilot** license + E3/E5 (or equivalent).
- **SPFx sample:** [Microsoft 365 Copilot API Explorer (SPFx)](https://adoption.microsoft.com/sample-solution-gallery/sample/pnp-sp-dev-spfx-web-parts-react-copilot-apis-explorer) — good reference for Chat (and other) APIs in a web part.

### Chat API (preview) — implementation hint

- **Docs:** [Chat API overview](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api/ai-services/chat/overview).
- **Flow:** Create or reuse a conversation; send user message (and optionally system context); get streamed or non-streamed response. In SPFx we use the same token as for Graph (with Copilot scope if required).
- **My Work Hub default prompt:** **“Summarize my pending work today.”** Build context string from Tasks, Approvals, and Recent (e.g. “The user has: 3 overdue To Do tasks (X, Y, Z), 2 Planner tasks due this week, 2 pending approvals (A, B), and 5 recommended files.”). Send as user message or context; show reply in Summary tab.

---

## 6. Permissions summary (SPFx package)

Request these in the SPFx package (and ensure admin consents if needed):

| Capability | Permissions |
|------------|-------------|
| To Do | `Tasks.Read`, `Tasks.ReadWrite` |
| Planner | `Tasks.Read`, `Tasks.ReadWrite`; optionally `Group.Read.All` |
| Approvals (beta) | `ApprovalSolution.Read`, `ApprovalSolutionResponse.ReadWrite` (or ReadWrite for solution) |
| Recent / files | `Files.Read.All` |
| Calendar (Meetings) | `Calendars.Read` |
| Copilot Chat | Copilot-related scope (see Copilot API docs; may be under `Copilot.Read` or similar — verify in [Copilot extensibility](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/)) |

Use **delegated** permissions so the web part acts as the signed-in user.

---

## 7. Implementation (completed)

1. **SPFx scaffold** — React, Fluent UI, MSGraphClient; configurable web part title in property pane.
2. **Tasks** — To Do (with linkedResources for flagged emails) + Planner; plan/list names; source icons; due badges (Today/Tomorrow); round checkbox to complete. **Tab 1.**
3. **Meetings** — Next 5 calendar events; date badges; organizer; time; link to open. **Tab 2.**
4. **Approvals** — List approvalItems (beta); client-side filter pending; markdown description; owner photo/name; reject-with-reason dialog. **Tab 3.**
5. **Recent** — `/me/drive/recent`; file-type icons; creator; last modified. **Tab 4.**
6. **Summary** — Copilot Chat API; rich context (tasks, approvals, recent); actionable prompt; ReactMarkdown response. **Tab 5.**
7. **Tabs** — Order: **Tasks | Meetings | Approvals | Recent | Summary.**

---

## 8. Links (quick reference)

| Topic | URL |
|-------|-----|
| SPFx web parts | https://learn.microsoft.com/en-us/training/modules/sharepoint-spfx-web-parts/ |
| To Do API | https://learn.microsoft.com/en-us/graph/todo-concept-overview |
| Planner API | https://learn.microsoft.com/en-us/graph/api/resources/planner-overview?view=graph-rest-1.0 |
| Planner tasks (me) | https://learn.microsoft.com/en-us/graph/api/planneruser-list-tasks?view=graph-rest-1.0 |
| Approvals app API | https://learn.microsoft.com/en-us/graph/approvals-app-api |
| approvalItem | https://learn.microsoft.com/en-us/graph/api/resources/approvalitem?view=graph-rest-beta |
| drive recent | https://learn.microsoft.com/en-us/graph/api/drive-recent?view=graph-rest-1.0 |
| Item insights | https://learn.microsoft.com/en-us/graph/item-insights-overview |
| Copilot APIs overview | https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/copilot-apis-overview |
| Graph auth | https://learn.microsoft.com/en-us/graph/auth/ |
