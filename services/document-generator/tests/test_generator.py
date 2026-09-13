from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

import pytest
from docx import Document

from timdoc_contracts import Customer, EquipmentItem, GenerationRequest, ServiceSettings
from timdoc_document_generator import DocumentGenerationError, DocumentGenerator, fill_blank


def _write_template(path: Path, title: str) -> None:
    document = Document()
    document.add_heading(title, level=1)
    document.add_paragraph("Хозяйство: {{ customer_name }} / ИНН {{ customer_inn }}")
    document.add_paragraph("Техника: {{ equipment_name }}")
    document.add_paragraph("Заводской номер: {{ serial_number }}")
    document.add_paragraph("Исполнитель: {{ service_center }} / {{ service_employee }}")
    document.add_paragraph("Дата: {{ document_date }}")
    document.add_paragraph("Адрес: {{ customer_address | blank(30) }}")
    document.add_paragraph("Контакт: {{ customer_contact | blank(20) }}")
    document.add_paragraph("Организация: {{ customer_name_inn | blank(10) }}")
    document.save(path)


@pytest.fixture
def template_directory(tmp_path: Path) -> Path:
    directory = tmp_path / "templates"
    directory.mkdir()
    _write_template(directory / "Талон.docx", "Талон")
    _write_template(directory / "Форма 4.docx", "Форма 4")
    _write_template(directory / "Форма 11.docx", "Форма 11")
    return directory


def _request(template_directory: Path, output_directory: Path) -> GenerationRequest:
    return GenerationRequest(
        customer=Customer(name="ООО «Рощинский»", inn="0268104130"),
        settings=ServiceSettings(employee_name="Абдрахманов Т.М."),
        equipment=[
            EquipmentItem(name="RSM SS-780-13", serial_number="M0S07013002106"),
            EquipmentItem(name="RSM SS-780-13", serial_number="M0S07013002105"),
        ],
        document_date="2026-09-07",
        template_directory=template_directory,
        output_directory=output_directory,
    )


def test_generates_three_documents_per_equipment_and_zip(
    tmp_path: Path, template_directory: Path
) -> None:
    before = {
        path.name: hashlib.sha256(path.read_bytes()).hexdigest()
        for path in template_directory.iterdir()
    }

    result = DocumentGenerator().generate(_request(template_directory, tmp_path / "output"))

    assert len(result.files) == 6
    assert Path(result.archive_path).is_file()
    assert {Path(path).name for path in result.files} == {
        "Талон.docx",
        "Форма 4.docx",
        "Форма 11.docx",
    }
    with zipfile.ZipFile(result.archive_path) as archive:
        assert len(archive.namelist()) == 6
        assert all(name.endswith(".docx") for name in archive.namelist())
    after = {
        path.name: hashlib.sha256(path.read_bytes()).hexdigest()
        for path in template_directory.iterdir()
    }
    assert after == before


def test_generated_document_contains_selected_customer_and_equipment(
    tmp_path: Path, template_directory: Path
) -> None:
    result = DocumentGenerator().generate(_request(template_directory, tmp_path / "output"))
    first = next(
        Path(path)
        for path in result.files
        if "M0S07013002106" in path and path.endswith("Талон.docx")
    )

    text = "\n".join(paragraph.text for paragraph in Document(first).paragraphs)
    assert "ООО «Рощинский»" in text
    assert "0268104130" in text
    assert "RSM SS-780-13" in text
    assert "M0S07013002106" in text
    assert "{{" not in text

    with zipfile.ZipFile(template_directory / "Талон.docx") as source_archive:
        source_parts = {
            name: source_archive.read(name)
            for name in source_archive.namelist()
            if name != "word/document.xml"
        }
    with zipfile.ZipFile(first) as result_archive:
        result_parts = {
            name: result_archive.read(name)
            for name in result_archive.namelist()
            if name != "word/document.xml"
        }
    assert result_parts == source_parts


def test_missing_template_is_a_hard_error(tmp_path: Path, template_directory: Path) -> None:
    (template_directory / "Форма 11.docx").unlink()

    with pytest.raises(DocumentGenerationError, match="Форма 11"):
        DocumentGenerator().generate(_request(template_directory, tmp_path / "output"))


def test_rejects_duplicate_serial_numbers(tmp_path: Path, template_directory: Path) -> None:
    request = _request(template_directory, tmp_path / "output")
    request.equipment.append(request.equipment[0])

    with pytest.raises(DocumentGenerationError, match="повторяется"):
        DocumentGenerator().generate(request)


def test_blank_lines_are_filled_or_left_for_handwriting(
    tmp_path: Path, template_directory: Path
) -> None:
    request = _request(template_directory, tmp_path / "output")
    request.customer.address = "с. Рощинский"
    request.customer.representative_position = "Главный инженер"
    request.customer.representative_name = "Иванов И.И."
    result = DocumentGenerator().generate(request)
    first = next(Path(path) for path in result.files if path.endswith("Форма 4.docx"))

    lines = {
        paragraph.text.split(":")[0]: paragraph.text for paragraph in Document(first).paragraphs
    }
    assert lines["Адрес"] == "Адрес: _________с. Рощинский_________"
    assert lines["Контакт"] == "Контакт: Главный инженер Иванов И.И."
    assert lines["Организация"] == "Организация: ООО «Рощинский», ИНН 0268104130"


def test_blank_filter_centers_values_and_keeps_empty_lines_blank() -> None:
    assert fill_blank("", 6) == "______"
    assert fill_blank(None, 3) == "___"
    assert fill_blank("ab", 6) == "__ab__"
    assert fill_blank("abc", 6) == "_abc__"
    assert fill_blank("  много   пробелов ", 20) == "___много пробелов___"
    assert fill_blank("слишком длинное значение", 5) == "слишком длинное значение"
