# Project Template

Use this as the starting point for new projects.

## Quick Start

1. Create a new project with `.\scripts\new-project.ps1 -Name "your-project-name"`.
2. Create a virtual environment.
3. Install the project and dev dependencies.
4. Copy `.env.example` to `.env`.
5. Run `python src/main.py`.
6. Run tests with `python -m pytest`.

```powershell
cd .\projects\your-project-name
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python src/main.py
python -m pytest
```

## Structure

- `src/` implementation
- `tests/` test suite
- `docs/` operational notes
- `config.example.yml` non-secret config template
- `.env.example` environment variable template

## Expectations

- Replace the placeholder README content with the project's real purpose,
  inputs, outputs, and run commands.
- Keep at least one smoke test that runs without private data so the scaffolded
  project is verifiable before pushing.
