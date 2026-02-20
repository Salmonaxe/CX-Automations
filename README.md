# CX-Automations

This repository is organized as a multi-project workspace.

## Structure

- `projects/` contains independent automation projects.
- `shared/` can hold reusable components used by multiple projects.

## Adding a New Project

Use the project generator script:

```powershell
.\scripts\new-project.ps1 -Name "my-first-project"
```

Then update the new project's `README.md` with its purpose and usage.

## Team Standard

Use `scripts/new-project.ps1` for all new project folders so structure stays consistent across the team.
See `CONTRIBUTING.md` for rules.

## Notes

Start with project-local code by default. Move code into `shared/` only when at least two projects need the same functionality.
