from pathlib import Path

UI = Path(__file__).parents[1] / "src" / "timdoc_app" / "ui"


def test_spa_is_local_accessible_and_contains_required_flow() -> None:
    html = (UI / "index.html").read_text(encoding="utf-8")
    assert "http://" not in html and "https://" not in html
    assert 'lang="ru"' in html
    assert 'aria-live="polite"' in html
    assert all(label in html for label in ("Выберите акт", "Проверьте данные", "Документы созданы"))


def test_ui_has_no_remote_assets() -> None:
    for path in UI.iterdir():
        if path.is_file():
            text = path.read_text(encoding="utf-8")
            assert "//cdn." not in text
