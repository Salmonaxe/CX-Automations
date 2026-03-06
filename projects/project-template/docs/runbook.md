# Runbook

## Local Setup

1. Create and activate a virtual environment.
2. Install the project with dev dependencies.
3. Copy `.env.example` to `.env` and set values.

```powershell
python -m pip install -e ".[dev]"
```

## Run

```powershell
python src/main.py
```

## Test

```powershell
python -m pytest
```
