import importlib.util
import subprocess
import sys
from pathlib import Path

import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
SCRIPT_CASES = [
    ("create_renewal_ppt.py", "Generate Cisco renewal opportunity PowerPoints"),
    ("create_renew_ops_ppt.py", "Generate Cisco renewal opportunity PowerPoints"),
    ("create_new_ops_ppt.py", "Generate Cisco new opportunities PowerPoint"),
]
MODULE_CASES = [
    ("renewal_baseline", "create_renewal_ppt.py"),
    ("renewal_enhanced", "create_renew_ops_ppt.py"),
    ("new_ops", "create_new_ops_ppt.py"),
]


def load_module(module_name: str, script_name: str):
    module_path = SRC_DIR / script_name
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


@pytest.mark.parametrize(("script_name", "expected_text"), SCRIPT_CASES)
def test_cli_help_runs(script_name, expected_text):
    result = subprocess.run(
        [sys.executable, str(SRC_DIR / script_name), "--help"],
        capture_output=True,
        text=True,
        cwd=PROJECT_ROOT,
        check=False,
    )
    assert result.returncode == 0
    assert expected_text in result.stdout


@pytest.mark.parametrize(("module_name", "script_name"), MODULE_CASES)
def test_template_validation_accepts_none(module_name, script_name):
    module = load_module(module_name, script_name)
    assert module.robust_check_template_file(None) is True


@pytest.mark.parametrize(("module_name", "script_name"), MODULE_CASES)
def test_template_validation_rejects_potx(module_name, script_name, capsys):
    module = load_module(module_name, script_name)
    assert module.robust_check_template_file("company-template.potx") is False
    captured = capsys.readouterr()
    assert ".pptx" in captured.err
