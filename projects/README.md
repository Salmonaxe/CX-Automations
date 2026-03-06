# Projects

Each folder in this directory is a standalone automation project.

## Current Contents

- `renewal-ppt-generator` is the active Cisco renewals and opportunities
  reporting project.
- `project-template` is the scaffold copied by `..\scripts\new-project.ps1`.

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

After scaffolding a Python project, install it with dev tooling before running
checks:

```powershell
cd .\projects\project-name
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest
```

## Recommended per-project layout

- `src/` implementation
- `tests/` tests
- `docs/` notes and runbooks
- `README.md` setup and usage
