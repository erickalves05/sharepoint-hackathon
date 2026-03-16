# My Work Hub — One productivity hub (SPFx web part)

**Workspace for maturing the hackathon idea:** a single SPFx web part that combines **Tasks/actions**, **My approvals**, and **My recent work**, with **Copilot** to add AI value. **Product name: My Work Hub.**

## What this folder is for

- **Discuss** scope, UX, and technical options.
- **Capture** decisions and references (Microsoft Learn, Graph, Copilot APIs).
- **Evolve** the concept before implementation.

## Quick concept

| Section | Purpose |
|--------|----------|
| **Tasks / actions hub** | One place for To Do tasks, Planner tasks, and (where possible) flagged Outlook items. |
| **My approvals hub** | Pending approval requests (e.g. Teams Approvals app / Power Automate) in one list with quick approve/reject. |
| **My recent work** | Recent/recommended files via **item insights** (Graph). |
| **Copilot spice** | Default: **“Summarize my pending work today”** (Chat API); optional follow-up. |

One web part, **four tabs** in order: **Tasks | Approvals | Recent | Summary** (Summary = Copilot).

## Documents in this workspace

| Document | Contents |
|----------|----------|
| [CONCEPT.md](./CONCEPT.md) | Unified product concept, user value, and high-level UX. |
| [POSSIBILITIES.md](./POSSIBILITIES.md) | Feature ideas, scope options, and what to build first vs later. |
| [TECHNICAL_NOTES.md](./TECHNICAL_NOTES.md) | APIs, permissions, and implementation notes from Microsoft Learn. |

## Next steps

1. Read CONCEPT.md and POSSIBILITIES.md.
2. Decide MVP scope (which sections + which Copilot feature first).
3. Use TECHNICAL_NOTES.md when scaffolding the SPFx project and calling Graph/Copilot.
