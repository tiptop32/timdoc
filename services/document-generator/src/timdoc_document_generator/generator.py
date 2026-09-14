from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import jinja2
from docxtpl import DocxTemplate  # type: ignore[import-untyped]

from timdoc_contracts import EquipmentItem, GenerationRequest, GenerationResult
from timdoc_document_generator import metrics


class DocumentGenerationError(ValueError):
    pass


BODY_FONT_SIZE = 22  # half-points: values are set in 11 pt regular


def blank_padding(
    value: object, width: int, size: int = BODY_FONT_SIZE, bold: bool = False
) -> tuple[int, int]:
    """How many underscores of the blank's own style go left and right of the value.

    The blank is `width` underscores of `size` half-points, bold or not; the value is set in
    11 pt regular. Padding is computed from Tahoma metrics so the line keeps its printed
    length, never growing (a longer line wraps and drags tab-aligned blocks after it).
    An empty value leaves the whole blank in place for handwriting.
    """
    text = " ".join(str(value or "").split())
    if not text:
        return width, 0
    unit = metrics.char_width(metrics.UNDERSCORE, bold=bold) * size / BODY_FONT_SIZE
    free = width * unit - metrics.text_width(text)
    count = int(free // unit) if free > 0 else 0
    return count // 2, count - count // 2


def blank_left(value: object, width: int, size: int = BODY_FONT_SIZE, bold: bool = False) -> str:
    return "_" * blank_padding(value, width, size, bold)[0]


def blank_right(value: object, width: int, size: int = BODY_FONT_SIZE, bold: bool = False) -> str:
    return "_" * blank_padding(value, width, size, bold)[1]


def template_environment() -> jinja2.Environment:
    environment = jinja2.Environment(autoescape=True)
    environment.filters["blank_left"] = blank_left
    environment.filters["blank_right"] = blank_right
    return environment


class DocumentGenerator:
    def generate(self, request: GenerationRequest) -> GenerationResult:
        templates = self._resolve_templates(request.template_directory)
        self._validate(request)
        environment = template_environment()
        request.output_directory.mkdir(parents=True, exist_ok=True)
        package_name = _available_name(
            request.output_directory,
            _safe_name(f"Комплект_{request.customer.name}_{request.document_date}"),
        )
        final_directory = request.output_directory / package_name
        final_archive = request.output_directory / f"{package_name}.zip"

        generated_relatives: list[Path] = []
        with tempfile.TemporaryDirectory(prefix=".timdoc-", dir=request.output_directory) as temp:
            work_root = Path(temp) / package_name
            for equipment in request.equipment:
                equipment_directory = work_root / _safe_name(equipment.serial_number)
                equipment_directory.mkdir(parents=True)
                context = self._context(request, equipment)
                for output_name, template_path in templates.items():
                    destination = equipment_directory / f"{output_name}.docx"
                    template = DocxTemplate(template_path)
                    undeclared = template.get_undeclared_template_variables(
                        jinja_env=environment, context=context
                    )
                    if undeclared:
                        names = ", ".join(sorted(undeclared))
                        raise DocumentGenerationError(
                            f"В шаблоне {template_path.name} не заполнены поля: {names}"
                        )
                    template.render(context, jinja_env=environment)
                    self._save_preserving_package(template, template_path, destination)
                    self._assert_rendered(destination)
                    generated_relatives.append(destination.relative_to(work_root))

            archive_work = Path(temp) / f"{package_name}.zip"
            with zipfile.ZipFile(archive_work, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                for relative in generated_relatives:
                    archive.write(work_root / relative, arcname=str(relative))
            if len(zipfile.ZipFile(archive_work).namelist()) != len(generated_relatives):
                raise DocumentGenerationError("ZIP-архив создан не полностью")

            shutil.move(str(work_root), final_directory)
            shutil.move(str(archive_work), final_archive)

        return GenerationResult(
            output_directory=str(final_directory),
            archive_path=str(final_archive),
            files=[str(final_directory / relative) for relative in generated_relatives],
        )

    @staticmethod
    def _resolve_templates(directory: Path) -> dict[str, Path]:
        required = {
            "Талон": directory / "Талон.docx",
            "Форма 4": directory / "Форма 4.docx",
            "Форма 11": directory / "Форма 11.docx",
        }
        missing = [name for name, path in required.items() if not path.is_file()]
        if missing:
            raise DocumentGenerationError(
                "Не найдены подготовленные шаблоны: " + ", ".join(missing)
            )
        return required

    @staticmethod
    def _validate(request: GenerationRequest) -> None:
        if not request.customer.name.strip():
            raise DocumentGenerationError("Укажите хозяйство")
        if not re.fullmatch(r"\d{10}|\d{12}", request.customer.inn):
            raise DocumentGenerationError("Укажите корректный ИНН хозяйства")
        if not request.equipment:
            raise DocumentGenerationError("В комплекте нет техники")
        serials = [item.serial_number.strip() for item in request.equipment]
        duplicate = next((serial for serial in serials if serials.count(serial) > 1), "")
        if duplicate:
            raise DocumentGenerationError(f"Заводской номер {duplicate} повторяется")
        if any(not serial for serial in serials):
            raise DocumentGenerationError("У каждой единицы должен быть заводской номер")

    @staticmethod
    def _context(request: GenerationRequest, equipment: EquipmentItem) -> dict[str, str]:
        customer = request.customer
        settings = request.settings
        formatted_date = _format_date(request.document_date)
        serial = equipment.serial_number.strip()
        context = {
            "customer_name": customer.name,
            "customer_inn": customer.inn,
            "customer_address": customer.address,
            "customer_country": customer.country,
            "customer_postal_code": customer.postal_code,
            "customer_region": customer.region,
            "customer_district": customer.district,
            "customer_locality": customer.locality,
            "customer_street": customer.street,
            "customer_house": customer.house,
            "customer_representative": customer.representative_name,
            "customer_representative_position": customer.representative_position,
            "customer_phone": customer.phone,
            "customer_email": customer.email,
            "customer_name_inn": f"{customer.name}, ИНН {customer.inn}",
            "customer_contact": " ".join(
                part
                for part in (
                    customer.representative_position,
                    customer.representative_name,
                    customer.phone,
                )
                if part.strip()
            ),
            "service_center": settings.organization,
            "service_employee_position": settings.employee_position,
            "service_employee": settings.employee_name,
            "service_location": settings.service_location,
            "equipment_name": equipment.name,
            "serial_number": serial,
            "manufacture_year": equipment.manufacture_year,
            "manufacturer": equipment.manufacturer,
            "document_date": formatted_date,
            "document_date_iso": request.document_date,
        }
        for index in range(14):
            context[f"serial_{index:02d}"] = serial[index] if index < len(serial) else ""
        return context

    @staticmethod
    def _assert_rendered(path: Path) -> None:
        with zipfile.ZipFile(path) as archive:
            document_xml = archive.read("word/document.xml")
        if b"{{" in document_xml or b"{%" in document_xml:
            raise DocumentGenerationError(f"В документе {path.name} остались маркеры шаблона")

    @staticmethod
    def _save_preserving_package(template: DocxTemplate, source: Path, destination: Path) -> None:
        """Replace only document.xml so unsupported Word package parts survive intact."""
        rendered_xml = template.docx._element.xml.encode("utf-8")
        try:
            with zipfile.ZipFile(source) as archive, zipfile.ZipFile(destination, "w") as output:
                for info in archive.infolist():
                    data = (
                        rendered_xml if info.filename == "word/document.xml" else archive.read(info)
                    )
                    output.writestr(info, data)
        except (KeyError, zipfile.BadZipFile, OSError) as error:
            raise DocumentGenerationError(
                f"Не удалось сохранить {destination.name}: {error}"
            ) from error


def _format_date(value: str) -> str:
    try:
        parsed = date.fromisoformat(value)
    except ValueError as error:
        raise DocumentGenerationError("Дата должна быть в формате ГГГГ-ММ-ДД") from error
    months = (
        "января",
        "февраля",
        "марта",
        "апреля",
        "мая",
        "июня",
        "июля",
        "августа",
        "сентября",
        "октября",
        "ноября",
        "декабря",
    )
    return f"«{parsed.day:02d}» {months[parsed.month - 1]} {parsed.year} г."


def _safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", "_", value).strip(" ._")
    return value[:120] or "Документы"


def _available_name(directory: Path, base: str) -> str:
    candidate = base
    index = 2
    while (directory / candidate).exists() or (directory / f"{candidate}.zip").exists():
        candidate = f"{base}_{index}"
        index += 1
    return candidate
