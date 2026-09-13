from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


class Serializable:
    def to_dict(self) -> dict[str, Any]:
        return asdict(self)  # type: ignore[call-overload, no-any-return]


@dataclass(slots=True)
class EquipmentItem(Serializable):
    name: str
    serial_number: str
    manufacture_year: str = ""
    manufacturer: str = ""


@dataclass(slots=True)
class ActData(Serializable):
    source_file: str
    document_date: str
    customer_name: str
    service_center_name: str
    equipment: list[EquipmentItem]
    customer_inn: str = ""
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class Customer(Serializable):
    name: str
    inn: str
    address: str = ""
    country: str = "РФ"
    postal_code: str = ""
    region: str = ""
    district: str = ""
    locality: str = ""
    street: str = ""
    house: str = ""
    representative_name: str = ""
    representative_position: str = ""
    phone: str = ""
    email: str = ""


@dataclass(slots=True)
class ServiceSettings(Serializable):
    organization: str = "ООО «Акрос РБ»"
    employee_position: str = "Инженер"
    employee_name: str = "Абдрахманов Т.М."
    service_location: str = ""
    template_directory: str = ""


@dataclass(slots=True)
class GenerationRequest(Serializable):
    customer: Customer
    settings: ServiceSettings
    equipment: list[EquipmentItem]
    document_date: str
    template_directory: Path
    output_directory: Path


@dataclass(slots=True)
class GenerationResult(Serializable):
    output_directory: str
    archive_path: str
    files: list[str]
