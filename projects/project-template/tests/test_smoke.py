from src.main import main


def test_smoke_runs(capsys):
    main()
    captured = capsys.readouterr()
    assert "project-template ran" in captured.out
