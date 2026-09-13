from __future__ import annotations

import sys
import types
from pathlib import Path
from typing import Any

import pytest

from timdoc_app import __main__ as entrypoint
from timdoc_app import desktop
from timdoc_app.desktop import DesktopBridge, run_desktop


class StubFacade:
    def __init__(self) -> None:
        self.calls: list[tuple[str, Any]] = []

    def bootstrap(self) -> dict[str, Any]:
        return {"settings": {}, "customers": [], "template_ready": False}

    def analyze_act(self, path: str) -> dict[str, Any]:
        if path.endswith(".txt"):
            raise ValueError("Выберите файл PDF")
        return {"act": {"source_file": path}}

    def prepare_templates(self, path: str) -> dict[str, Any]:
        self.calls.append(("prepare_templates", path))
        return {"template_ready": True}

    def save_customer(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload

    def save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload

    def generate(self, payload: dict[str, Any]) -> dict[str, Any]:
        if not payload.get("output_directory"):
            raise OSError("Нет доступа к папке")
        return {"files": ["a", "b", "c"]}


class FakeWindow:
    def __init__(self, selection: list[str] | None) -> None:
        self.selection = selection
        self.dialogs: list[tuple[Any, dict[str, Any]]] = []

    def create_file_dialog(self, kind: Any, **options: Any) -> list[str] | None:
        self.dialogs.append((kind, options))
        return self.selection


@pytest.fixture
def fake_webview(monkeypatch: pytest.MonkeyPatch) -> types.ModuleType:
    module = types.ModuleType("webview")
    module.OPEN_DIALOG = "open"  # type: ignore[attr-defined]
    module.FOLDER_DIALOG = "folder"  # type: ignore[attr-defined]
    created: dict[str, Any] = {}

    def create_window(title: str, url: str, **options: Any) -> FakeWindow:
        created.update({"title": title, "url": url, **options})
        return FakeWindow(["/tmp/chosen.pdf"])

    def start(**options: Any) -> None:
        created["started"] = options

    module.create_window = create_window  # type: ignore[attr-defined]
    module.start = start  # type: ignore[attr-defined]
    module.created = created  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "webview", module)
    return module


def test_bridge_wraps_results_and_errors_into_envelopes() -> None:
    bridge = DesktopBridge(StubFacade())  # type: ignore[arg-type]

    assert bridge.bootstrap() == {
        "ok": True,
        "data": {"settings": {}, "customers": [], "template_ready": False},
    }
    assert bridge.analyze_act("act.pdf") == {
        "ok": True,
        "data": {"act": {"source_file": "act.pdf"}},
    }
    assert bridge.analyze_act("act.txt") == {"ok": False, "error": "Выберите файл PDF"}
    assert bridge.prepare_templates("/templates")["data"] == {"template_ready": True}
    assert bridge.save_customer({"name": "ООО «Рощинский»"})["data"]["name"] == "ООО «Рощинский»"
    assert bridge.save_settings({"employee_name": "Абдрахманов Т.М."})["ok"] is True
    assert bridge.generate({"output_directory": "/out"})["data"]["files"] == ["a", "b", "c"]
    assert bridge.generate({}) == {"ok": False, "error": "Нет доступа к папке"}


def test_bridge_file_dialogs_return_the_selected_path(fake_webview: types.ModuleType) -> None:
    bridge = DesktopBridge(StubFacade())  # type: ignore[arg-type]
    bridge.window = FakeWindow(["/data/Акт.pdf"])
    assert bridge.choose_pdf() == {"ok": True, "path": "/data/Акт.pdf"}
    assert bridge.window.dialogs[0][0] == "open"
    assert bridge.window.dialogs[0][1]["file_types"] == ("PDF (*.pdf)",)

    bridge.window = FakeWindow(None)
    assert bridge.choose_pdf() == {"ok": True, "path": ""}
    assert bridge.choose_folder() == {"ok": True, "path": ""}

    bridge.window = FakeWindow(["/data/out"])
    assert bridge.choose_folder() == {"ok": True, "path": "/data/out"}
    assert bridge.window.dialogs[0][0] == "folder"


def test_run_desktop_opens_the_local_spa(
    fake_webview: types.ModuleType, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(desktop, "create_default_facade", lambda: StubFacade())
    run_desktop()
    created = fake_webview.created  # type: ignore[attr-defined]
    assert created["title"] == "Timdoc"
    assert created["url"].startswith("file://") and created["url"].endswith("/ui/index.html")
    assert Path(created["url"].removeprefix("file://")).is_file()
    assert isinstance(created["js_api"], DesktopBridge)
    assert created["js_api"].window is not None
    assert created["started"] == {"debug": False, "private_mode": True}


def test_entrypoint_self_test_bootstraps_without_a_window(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    facade = StubFacade()
    monkeypatch.setattr("timdoc_app.facade.create_default_facade", lambda: facade)
    monkeypatch.setattr(sys, "argv", ["timdoc", "--self-test"])
    entrypoint.main()
    assert "Timdoc self-test: OK" in capsys.readouterr().out


def test_entrypoint_starts_the_desktop_app_by_default(monkeypatch: pytest.MonkeyPatch) -> None:
    launched: list[bool] = []
    monkeypatch.setattr(entrypoint, "run_desktop", lambda: launched.append(True))
    monkeypatch.setattr(sys, "argv", ["timdoc"])
    entrypoint.main()
    assert launched == [True]
