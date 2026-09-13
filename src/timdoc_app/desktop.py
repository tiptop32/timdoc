from __future__ import annotations

from pathlib import Path
from typing import Any

from timdoc_app.facade import AppFacade, create_default_facade


class DesktopBridge:
    def __init__(self, facade: AppFacade) -> None:
        self.facade = facade
        self.window: Any = None

    def bootstrap(self) -> dict[str, Any]:
        return self._call(self.facade.bootstrap)

    def analyze_act(self, path: str) -> dict[str, Any]:
        return self._call(self.facade.analyze_act, path)

    def prepare_templates(self, path: str) -> dict[str, Any]:
        return self._call(self.facade.prepare_templates, path)

    def save_customer(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self._call(self.facade.save_customer, payload)

    def save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self._call(self.facade.save_settings, payload)

    def generate(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self._call(self.facade.generate, payload)

    def choose_pdf(self) -> dict[str, Any]:
        import webview  # type: ignore[import-not-found]

        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("PDF (*.pdf)",),
        )
        return {"ok": True, "path": result[0] if result else ""}

    def choose_folder(self) -> dict[str, Any]:
        import webview  # type: ignore[import-not-found]

        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        return {"ok": True, "path": result[0] if result else ""}

    @staticmethod
    def _call(function: Any, *args: Any) -> dict[str, Any]:
        try:
            return {"ok": True, "data": function(*args)}
        except (OSError, ValueError) as error:
            return {"ok": False, "error": str(error)}


def run_desktop() -> None:
    import webview  # type: ignore[import-not-found]

    bridge = DesktopBridge(create_default_facade())
    page = Path(__file__).with_name("ui") / "index.html"
    bridge.window = webview.create_window(
        "Timdoc",
        page.as_uri(),
        js_api=bridge,
        width=1180,
        height=800,
        min_size=(920, 680),
    )
    webview.start(debug=False, private_mode=True)
