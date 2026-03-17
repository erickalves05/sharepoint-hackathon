# My Work Hub — One productivity hub (SPFx web part)

**Implemented:** A single SPFx web part that combines **Tasks**, **Meetings**, **Approvals**, **Recent files**, and **Copilot** in one place. **Product name: My Work Hub.**

## What this folder is for

- **Discuss** scope, UX, and technical options.
- **Capture** decisions and references (Microsoft Learn, Graph, Copilot APIs).
- **Document** implementation progress and hackathon submission.

## Implemented features

| Tab | Purpose |
|-----|---------|
| **Tasks** | Unified To Do + Planner tasks, including flagged emails. Source icons (Outlook, To Do, Planner), due date badges (Today/Tomorrow), plan/list names, round checkbox to complete. |
| **Meetings** | Next 5 upcoming calendar events. Date badges (Today/Tomorrow), organizer, time range, link to open meeting. |
| **Approvals** | Pending approvals with markdown descriptions, requester photo/name, reject-with-reason dialog. |
| **Recent** | Recent OneDrive files via `/me/drive/recent`. File-type icons (Excel, Word, PowerPoint, PDF), creator, last modified. |
| **Summary** | Copilot summary with rich context (tasks, approvals, recent files) and actionable recommendations. |

One web part, **five tabs** in order: **Tasks | Meetings | Approvals | Recent | Summary**.

## Documents in this workspace

| Document | Contents |
|----------|----------|
| [CONCEPT.md](./CONCEPT.md) | Unified product concept, user value, and high-level UX. |
| [POSSIBILITIES.md](./POSSIBILITIES.md) | Feature ideas, scope options, and what to build first vs later. |
| [TECHNICAL_NOTES.md](./TECHNICAL_NOTES.md) | APIs, permissions, and implementation notes (updated with actual choices). |
| [SUBMISSION.md](./SUBMISSION.md) | Hackathon submission form draft (title, name, description). |
