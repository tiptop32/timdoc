from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from docx import Document

from timdoc_app.facade import AppFacade, application_data_directory
from timdoc_contracts import ActData, EquipmentItem


class StubParser:
    def parse(self, path: Path) -> ActData:
        return ActData(
            str(path),
            "2026-09-07",
            "ООО «Рощинский»",
            "ООО «Акрос РБ»",
            [EquipmentItem("RSM SS-780-13", "M0S07013002106")],
        )


class StubPreparer:
    def prepare(self, source: Path, destination: Path) -> Path:
        destination.mkdir(parents=True, exist_ok=True)
        for name in ("Талон.docx", "Форма 4.docx", "Форма 11.docx"):
            Document().save(destination / name)
        return destination


def test_facade_matches_customer_and_persists_defaults(tmp_path: Path) -> None:
    facade = AppFacade(tmp_path, parser=StubParser(), preparer=StubPreparer())  # type: ignore[arg-type]
    facade.save_customer({"name": "ООО Рощинский", "inn": "0268104130", "phone": "+7 900"})
    result = facade.analyze_act("act.pdf")
    assert result["customer"]["inn"] == "0268104130"
    assert facade.bootstrap()["settings"]["organization"] == "ООО «Акрос РБ»"


def test_facade_generates_three_documents_and_writes_trace(tmp_path: Path) -> None:
    source = tmp_path / "source"
    output = tmp_path / "output"
    source.mkdir()
    facade = AppFacade(tmp_path / "data", preparer=StubPreparer())  # type: ignore[arg-type]
    result = facade.generate(
        {
            "customer": {"name": "ООО «Рощинский»", "inn": "0268104130"},
            "settings": {
                "organization": "ООО «Акрос РБ»",
                "employee_position": "Инженер",
                "employee_name": "Абдрахманов Т.М.",
                "template_directory": str(source),
            },
            "equipment": [{"name": "RSM SS-780-13", "serial_number": "M0S07013002106"}],
            "document_date": "2026-09-07",
            "output_directory": str(output),
        }
    )
    assert len(result["files"]) == 3
    event = json.loads((tmp_path / "data" / "events.jsonl").read_text(encoding="utf-8"))
    assert event["event"] == "generation_completed"
    assert event["document_count"] == 3


def test_facade_validates_output_and_template_directories(tmp_path: Path) -> None:
    facade = AppFacade(tmp_path / "data", preparer=StubPreparer())  # type: ignore[arg-type]
    payload = {
        "customer": {"name": "ООО «Рощинский»", "inn": "0268104130"},
        "settings": {"template_directory": str(tmp_path)},
        "equipment": [{"name": "RSM SS-780-13", "serial_number": "M0S07013002106"}],
        "document_date": "2026-09-07",
        "output_directory": "",
    }
    with pytest.raises(ValueError, match="папку сохранения"):
        facade.generate(payload)

    payload["output_directory"] = str(tmp_path / "out")
    payload["settings"] = {"template_directory": "  "}
    with pytest.raises(ValueError, match="шаблонов"):
        facade.generate(payload)


def test_facade_prepare_templates_and_settings_round_trip(tmp_path: Path) -> None:
    facade = AppFacade(tmp_path / "data", parser=StubParser(), preparer=StubPreparer())  # type: ignore[arg-type]
    assert facade.bootstrap()["template_ready"] is False

    result = facade.prepare_templates(str(tmp_path))
    assert result["template_ready"] is True
    assert result["settings"]["template_directory"] == str(tmp_path)
    assert facade.bootstrap()["template_ready"] is True

    saved = facade.save_settings({"employee_name": "Иванов И.И.", "unknown_field": "x"})
    assert saved["employee_name"] == "Иванов И.И."
    assert facade.bootstrap()["settings"]["employee_name"] == "Иванов И.И."
    assert facade.analyze_act("act.pdf")["customer"]["name"] == "ООО «Рощинский»"


def test_application_data_directory_follows_the_platform(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setenv("LOCALAPPDATA", r"C:\Users\tim\AppData\Local")
    assert application_data_directory() == Path(r"C:\Users\tim\AppData\Local") / "Timdoc"

    monkeypatch.setattr(sys, "platform", "darwin")
    assert (
        application_data_directory() == Path.home() / "Library" / "Application Support" / "Timdoc"
    )

    monkeypatch.setattr(sys, "platform", "linux")
    monkeypatch.setenv("XDG_DATA_HOME", "/srv/data")
    assert application_data_directory() == Path("/srv/data") / "Timdoc"
