from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from pypdf import PdfReader
from pypdf.errors import PdfReadError

from timdoc_contracts import ActData, EquipmentItem


class ActParseError(ValueError):
    pass


class ActParser:
    def parse(self, path: Path) -> ActData:
        path = Path(path)
        if path.suffix.casefold() != ".pdf":
            raise ActParseError("Выберите файл PDF")
        if not path.is_file():
            raise ActParseError(f"PDF не найден: {path}")

        try:
            reader = PdfReader(path)
            if reader.is_encrypted and not reader.decrypt(""):
                raise ActParseError("PDF защищён паролем")
            pages = [self._extract_page(page) for page in reader.pages]
        except (PdfReadError, OSError, ValueError) as error:
            if isinstance(error, ActParseError):
                raise
            raise ActParseError(f"Не удалось прочитать PDF: {error}") from error

        text = "\n".join(pages).strip()
        if len(text) < 40:
            raise ActParseError("В PDF нет текстового слоя. OCR пока не поддерживается")

        warnings: list[str] = []
        date = self._parse_date(text)
        customer = self._parse_party(text, "Покупатель")
        service_center = self._parse_party(text, "Продавец")
        items = self._parse_equipment(text)
        inn = self._parse_customer_inn(text, customer)

        if not date:
            warnings.append("Не найдена дата акта")
        if not customer:
            warnings.append("Не найдено хозяйство-покупатель")
        if not service_center:
            warnings.append("Не найден продавец или сервисный центр")
        if not items:
            warnings.append("Не найдены техника и заводские номера")
        if not inn:
            warnings.append("ИНН покупателя отсутствует в акте, выберите хозяйство или введите ИНН")

        return ActData(
            source_file=str(path.resolve()),
            document_date=date,
            customer_name=customer,
            service_center_name=service_center,
            equipment=items,
            customer_inn=inn,
            warnings=warnings,
        )

    @staticmethod
    def _extract_page(page: Any) -> str:
        extract_text = page.extract_text
        try:
            return str(extract_text(extraction_mode="layout") or "")
        except TypeError:
            return str(extract_text() or "")

    @staticmethod
    def _parse_party(text: str, role: str) -> str:
        pattern = re.compile(
            rf"\((ООО\s+[«\"]?[^»\"\n,)]+[»\"]?)\),\s*именуемое\s+в\s+дальнейшем\s+[«\"]{role}[»\"]",
            re.IGNORECASE,
        )
        match = pattern.search(text)
        if match:
            return _clean_party(match.group(1))

        role_position = text.casefold().find(role.casefold())
        prefix = text[max(0, role_position - 300) : role_position] if role_position >= 0 else text
        candidates = re.findall(r"ООО\s+[«\"][^»\"\n]+[»\"]", prefix, re.IGNORECASE)
        return _clean_party(candidates[-1]) if candidates else ""

    @staticmethod
    def _parse_date(text: str) -> str:
        months = {
            "января": 1,
            "февраля": 2,
            "марта": 3,
            "апреля": 4,
            "мая": 5,
            "июня": 6,
            "июля": 7,
            "августа": 8,
            "сентября": 9,
            "октября": 10,
            "ноября": 11,
            "декабря": 12,
        }
        match = re.search(
            r"[«\"]?(\d{1,2})[»\"]?\s+(" + "|".join(months) + r")\s+(20\d{2})\s*г?\.?,?",
            text,
            re.IGNORECASE,
        )
        if not match:
            return ""
        day, month_name, year = match.groups()
        return f"{year}-{months[month_name.casefold()]:02d}-{int(day):02d}"

    @staticmethod
    def _parse_customer_inn(text: str, customer: str) -> str:
        if customer:
            position = text.find(customer)
            if position >= 0:
                nearby = text[position : position + 250]
                match = re.search(r"\b(?:ИНН|инн)\s*[:№]?\s*(\d{10}|\d{12})\b", nearby)
                if match:
                    return match.group(1)
        matches = re.findall(r"\b(?:ИНН|инн)\s*[:№]?\s*(\d{10}|\d{12})\b", text)
        return matches[0] if len(set(matches)) == 1 else ""

    @staticmethod
    def _parse_equipment(text: str) -> list[EquipmentItem]:
        table_match = re.search(
            r"(?:следующий\s+Товар|следующую\s+Продукцию)\s*:(.*?)(?:\n\s*2\.\s|\n\s*Товар\s+передан)",
            text,
            re.IGNORECASE | re.DOTALL,
        )
        table = table_match.group(1) if table_match else text
        lines = [_space(line) for line in table.splitlines()]
        serial_pattern = re.compile(
            r"\b(?=[A-ZА-Я0-9-]{11,24}\b)(?=[A-ZА-Я0-9-]*\d)[A-ZА-Я0-9-]+\b"
        )
        model_pattern = re.compile(r"\bRSM\s+[A-Z0-9]+(?:-[A-Z0-9]+)+\b", re.IGNORECASE)
        result: list[EquipmentItem] = []
        seen: set[str] = set()

        for index, line in enumerate(lines):
            candidates = [
                value for value in serial_pattern.findall(line) if not value.startswith("20")
            ]
            if not candidates:
                continue
            serial = candidates[0]
            if serial in seen:
                continue
            context = " ".join(_row_context(lines, index, serial_pattern))
            model_match = model_pattern.search(context)
            if model_match:
                name = _space(model_match.group(0)).upper()
            else:
                name = _fallback_product_name(lines, index, serial)
            year_match = re.search(r"\b(20\d{2})\b", line)
            manufacturer_match = re.search(r"(?:АО|ООО|ПАО)\s+[«\"][^»\"]+[»\"]", line)
            result.append(
                EquipmentItem(
                    name=name,
                    serial_number=serial,
                    manufacture_year=year_match.group(1) if year_match else "",
                    manufacturer=manufacturer_match.group(0) if manufacturer_match else "",
                )
            )
            seen.add(serial)
        return result


def _space(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\xa0", " ")).strip()


def _clean_party(value: str) -> str:
    value = _space(value)
    match = re.match(r"ООО\s+[«\"]?(.+?)[»\"]?$", value, re.IGNORECASE)
    return f"ООО «{match.group(1).strip()}»" if match else value


def _row_context(lines: list[str], index: int, serial_pattern: re.Pattern[str]) -> list[str]:
    """Lines of one table row: the serial line plus wrapped neighbours up to the next serial.

    A row of the act may wrap onto several PDF lines, so the model is searched around the
    serial number, but never across a line that carries another serial number: that line
    already belongs to a different piece of equipment.
    """
    start = index
    for candidate in range(index - 1, max(-1, index - 3), -1):
        if serial_pattern.search(lines[candidate]):
            break
        start = candidate
    end = index + 1
    for candidate in range(index + 1, min(len(lines), index + 3)):
        if serial_pattern.search(lines[candidate]):
            break
        end = candidate + 1
    return lines[start:end]


def _fallback_product_name(lines: list[str], index: int, serial: str) -> str:
    fragments: list[str] = []
    for line in lines[max(0, index - 1) : index + 2]:
        cleaned = _space(line.replace(serial, ""))
        cleaned = re.sub(r"^\d+\s+", "", cleaned)
        cleaned = re.sub(r"\b20\d{2}\b.*$", "", cleaned).strip()
        if cleaned and not re.fullmatch(r"\d+", cleaned):
            fragments.append(cleaned)
    return _space(" ".join(fragments))
