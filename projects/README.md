# Projects

Each folder in this directory is a standalone automation project.

## Naming

Use clear, task-based names, for example:

- `invoice-reconciliation`
- `zendesk-ticket-sync`
- `daily-kpi-report`

## Standard Creation

Create projects with:

```powershell
.\scripts\new-project.ps1 -Name "project-name"
```

## Recommended per-project layout

- `src/` implementation
- `tests/` tests
- `docs/` notes and runbooks
- `README.md` setup and usage
