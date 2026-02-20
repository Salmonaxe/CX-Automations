# Contributing

## Standard Project Creation

Create all new automation projects using the repository script:

```powershell
.\scripts\new-project.ps1 -Name "your-project-name"
```

Use lowercase letters, numbers, and hyphens only.

Examples:

- `invoice-reconciliation`
- `zendesk-ticket-sync`
- `daily-kpi-report`

Do not manually create project folders unless there is a documented exception.

## Project Rules

- Keep implementation in `src/`, tests in `tests/`, and operational notes in `docs/`.
- Keep dependencies scoped to each project.
- Keep secrets out of source control; use `.env` locally and commit only `.env.example`.
- Move code to `shared/` only after real reuse across multiple projects.

## Overwriting an Existing Project Folder

If you intentionally want to recreate an existing project folder from template:

```powershell
.\scripts\new-project.ps1 -Name "your-project-name" -Force
```
