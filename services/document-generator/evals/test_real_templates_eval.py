"""Приемка на настоящих шаблонах знакомого и настоящем Word.

Шаблоны лежат вне Git (папка doc_templates), поэтому eval пропускается там, где их нет.
На macOS документы дополнительно открываются в Word с включенными предупреждениями:
если Word покажет диалог восстановления, открытие не завершится и eval упадет по таймауту.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import pytest
from lxml import etree as ET
from pypdf import PdfReader

from timdoc_contracts import Customer, EquipmentItem, GenerationRequest, ServiceSettings
from timdoc_document_generator import DocumentGenerator, TemplatePreparer
from timdoc_document_generator.templates import OFFICE_GROUP_CONTAINER, W

TEMPLATES = Path(
    os.environ.get("TIMDOC_TEMPLATES", "/Users/tiptop32/my_git_reps/timdoc/doc_templates")
)


def _word_container_writable() -> bool:
    """Word for Mac converts only inside its group container; a sandbox may block it."""
    root = OFFICE_GROUP_CONTAINER.expanduser()
    if not root.is_dir():
        return False
    try:
        with tempfile.TemporaryDirectory(prefix="probe-", dir=root):
            return True
    except OSError:
        return False


WORD_AVAILABLE = sys.platform == "win32" or (
    sys.platform == "darwin" and _word_container_writable()
)

# Открывает копию документа внутри контейнера Word (без диалогов доступа), с включенными
# предупреждениями, печатает в PDF и закрывает. Диалог восстановления блокирует "open".
WORD_RENDER_SCRIPT = """\
on run argv
  tell application "Microsoft Word"
    set display alerts to alerts all
    with timeout of 30 seconds
      open file name (item 1 of argv)
    end timeout
    set display alerts to alerts none
    save as active document file name (item 2 of argv) file format format PDF
    close active document saving no
  end tell
end run
"""


def _render_with_word(document: Path) -> int:
    """Return the page count Word produces for the document; raise if Word cannot open it."""
    staging_root = OFFICE_GROUP_CONTAINER.expanduser() / "Timdoc"
    staging_root.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="render-", dir=staging_root) as temp:
        staged = Path(temp) / document.name
        pdf = Path(temp) / (document.stem + ".pdf")
        shutil.copyfile(document, staged)
        completed = subprocess.run(
            ["osascript", "-e", WORD_RENDER_SCRIPT, str(staged), str(pdf)],
            capture_output=True,
            text=True,
            timeout=90,
            check=False,
        )
        assert completed.returncode == 0, f"Word не открыл {document.name}: {completed.stderr}"
        return len(PdfReader(pdf).pages)


def _text(path: Path) -> list[str]:
    """Text of every paragraph, including those inside tables and frames; tabs kept."""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    lines = []
    for paragraph in root.iter(f"{{{W}}}p"):
        parts = []
        for node in paragraph.iter():
            if node.tag == f"{{{W}}}t":
                parts.append(node.text or "")
            elif node.tag == f"{{{W}}}tab":
                parts.append("\t")
        lines.append("".join(parts))
    return lines


@pytest.mark.eval
@pytest.mark.skipif(not TEMPLATES.is_dir(), reason="real templates live outside Git")
@pytest.mark.skipif(
    not WORD_AVAILABLE, reason="needs Microsoft Word and write access to its container"
)
def test_real_templates_fill_like_the_reference_documents(tmp_path: Path) -> None:
    prepared = TemplatePreparer().prepare(TEMPLATES, tmp_path / "prepared")
    request = GenerationRequest(
        customer=Customer(
            name="ООО «Рощинский»",
            inn="0278000000",
            address="453100, Республика Башкортостан, с. Рощинский, ул. Ленина, 1",
            postal_code="453100",
            region="Республика Башкортостан",
            district="Стерлитамакский район",
            locality="с. Рощинский",
            street="ул. Ленина",
            house="1",
            representative_name="Иванов И.И.",
            representative_position="Главный инженер",
            phone="+7 917 000-00-00",
            email="info@roshinsky.ru",
        ),
        settings=ServiceSettings(service_location="г. Уфа"),
        equipment=[EquipmentItem(name="RSM SS-780-13", serial_number="M0S07013002106")],
        document_date="2026-09-07",
        template_directory=prepared,
        output_directory=tmp_path / "output",
    )
    result = DocumentGenerator().generate(request)
    files = {Path(path).name: Path(path) for path in result.files}
    assert set(files) == {"Талон.docx", "Форма 4.docx", "Форма 11.docx"}

    for name in ("Форма 4.docx", "Форма 11.docx"):
        lines = _text(files[name])
        model = next(line for line in lines if line.startswith("провел работы по обслуживанию"))
        serial = next(line for line in lines if line.startswith("заводской номер №"))
        customer = next(line for line in lines if line.startswith("совместно с Потребителем"))
        assert "RSM SS-780-13" in model and model.count("_") > 20
        assert "M0S07013002106" in serial and serial.count("_") > 20
        assert "ООО «Рощинский», ИНН 0278000000" in customer
        assert any("г. Уфа" in line for line in lines)
        assert any("«07» сентября 2026 г." in line for line in lines)
        assert any("Главный инженер Иванов И.И. +7 917 000-00-00" in line for line in lines)
        position = next(line for line in lines if "Инженер" in line and "\t" in line)
        assert position.endswith("_" * 35), "линия должности потребителя потеряна"
        signature = next(line for line in lines if "Абдрахманов Т.М." in line)
        assert signature.endswith("________     _____________________")
        assert not any("{{" in line for line in lines)

    ticket_text = "\n".join(_text(files["Талон.docx"]))
    assert ticket_text.count("_____________________ Абдрахманов Т.М. М.П.") == 2
    assert ticket_text.count("ООО «Рощинский»") == 2
    assert ticket_text.count("«07» сентября 2026 г.") == 2

    if sys.platform == "darwin":
        for name, path in files.items():
            assert _render_with_word(path) == 1, f"{name} не поместился на одну страницу"
