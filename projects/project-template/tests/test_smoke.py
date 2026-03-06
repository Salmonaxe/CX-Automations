import importlib.util
from pathlib import Path


def load_main():
    project_root = Path(__file__).resolve().parents[1]
    module_path = project_root / "src" / "main.py"
    spec = importlib.util.spec_from_file_location("project_template_main", module_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module.main


def test_smoke_runs(capsys):
    main = load_main()
    main()
    captured = capsys.readouterr()
    assert "project-template ran" in captured.out
