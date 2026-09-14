from __future__ import annotations

import json
import os
import sys
from dataclasses import fields
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from timdoc_act_parser import ActParser
from timdoc_contracts import Customer, EquipmentItem, GenerationRequest, ServiceSettings
from timdoc_customer_store import CustomerStore
from timdoc_document_generator import DocumentGenerator, TemplatePreparer


def application_data_directory() -> Path:
    if sys.platform == "win32":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    else:
        root = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    return root / "Timdoc"


class AppFacade:
    def __init__(
        self,
        data_directory: Path,
        *,
        parser: ActParser | None = None,
        store: CustomerStore | None = None,
        preparer: TemplatePreparer | None = None,
        generator: DocumentGenerator | None = None,
    ) -> None:
        self.data_directory = Path(data_directory)
        self.data_directory.mkdir(parents=True, exist_ok=True)
        self.parser = parser or ActParser()
        self.store = store or CustomerStore(self.data_directory / "timdoc.sqlite3")
        self.preparer = preparer or TemplatePreparer()
        self.generator = generator or DocumentGenerator()
        self.prepared_templates = self.data_directory / "prepared-templates"

    def bootstrap(self) -> dict[str, Any]:
        settings = self.store.get_settings()
        return {
            "settings": settings.to_dict(),
            "customers": [customer.to_dict() for customer in self.store.list_customers()],
            "template_ready": self._templates_ready(),
            "data_directory": str(self.data_directory),
        }

    def analyze_act(self, path: str) -> dict[str, Any]:
        act = self.parser.parse(Path(path))
        customer = self.store.find_customer(name=act.customer_name, inn=act.customer_inn)
        if customer is None:
            customer = Customer(name=act.customer_name, inn=act.customer_inn)
        return {"act": act.to_dict(), "customer": customer.to_dict()}

    def save_customer(self, payload: dict[str, Any]) -> dict[str, Any]:
        customer = _from_payload(Customer, payload)
        return self.store.upsert_customer(customer).to_dict()

    def save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        settings = _from_payload(ServiceSettings, payload)
        return self.store.save_settings(settings).to_dict()

    def prepare_templates(self, source_directory: str) -> dict[str, Any]:
        settings = self.store.get_settings()
        settings.template_directory = source_directory
        self.store.save_settings(settings)
        self.preparer.prepare(Path(source_directory), self.prepared_templates)
        return {"template_ready": True, "settings": settings.to_dict()}

    def generate(self, payload: dict[str, Any]) -> dict[str, Any]:
        customer = _from_payload(Customer, payload.get("customer", {}))
        settings = _from_payload(ServiceSettings, payload.get("settings", {}))
        equipment = [_from_payload(EquipmentItem, item) for item in payload.get("equipment", [])]
        self.store.upsert_customer(customer)
        self.store.save_settings(settings)
        self.preparer.prepare(Path(settings.template_directory), self.prepared_templates)
        output_directory = str(payload.get("output_directory", "")).strip()
        if not output_directory:
            raise ValueError("Выберите папку сохранения")
        if not settings.template_directory.strip():
            raise ValueError("Выберите папку исходных шаблонов в настройках")
        request = GenerationRequest(
            customer=customer,
            settings=settings,
            equipment=equipment,
            document_date=str(payload.get("document_date", "")),
            template_directory=self.prepared_templates,
            output_directory=Path(output_directory),
        )
        result = self.generator.generate(request)
        self._write_event(
            "generation_completed",
            customer=customer.name,
            equipment_count=len(equipment),
            document_count=len(result.files),
            archive_path=result.archive_path,
        )
        return result.to_dict()

    def _templates_ready(self) -> bool:
        return all(
            (self.prepared_templates / name).is_file()
            for name in ("Талон.docx", "Форма 4.docx", "Форма 11.docx")
        )

    def _write_event(self, event: str, **details: Any) -> None:
        record = {
            "timestamp": datetime.now(UTC).isoformat(),
            "event": event,
            **details,
        }
        with (self.data_directory / "events.jsonl").open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(record, ensure_ascii=False) + "\n")


def create_default_facade() -> AppFacade:
    return AppFacade(application_data_directory())


def _from_payload(model: type[Any], payload: dict[str, Any]) -> Any:
    allowed = {item.name for item in fields(model)}
    return model(**{key: value for key, value in payload.items() if key in allowed})
