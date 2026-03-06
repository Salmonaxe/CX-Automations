# CX-Automations

This repository is organized as a multi-project workspace.

## Structure

- `projects/` contains independent automation projects.
- `shared/` can hold reusable components used by multiple projects.
- `Archive/` contains older one-off scripts kept for reference while active work
  stays under `projects/`.

## Current Projects

- `projects/renewal-ppt-generator/` contains the active Cisco renewals and new
  opportunities automation.
- `projects/project-template/` is the scaffold used by
  `scripts/new-project.ps1`.

## Adding a New Project

Use the project generator script:

```powershell
.\scripts\new-project.ps1 -Name "my-first-project"
```

Then update the new project's `README.md` with its purpose and usage.

## Team Standard

Use `scripts/new-project.ps1` for all new project folders so structure stays consistent across the team.
See `CONTRIBUTING.md` for rules.

## Verification

Run checks from the project you are changing. For Python projects in this repo,
install the project with its dev tools first:

```powershell
cd .\projects\<project-name>
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest
```

## Notes

Start with project-local code by default. Move code into `shared/` only when at least two projects need the same functionality.
