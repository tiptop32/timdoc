from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any, Protocol

from lxml import etree as ET  # type: ignore[import-untyped]


class TemplatePreparationError(ValueError):
    pass


class DocumentConverter(Protocol):
    def convert(self, source: Path, destination: Path) -> None: ...


def default_converter(platform: str = sys.platform) -> DocumentConverter:
    """Pick the DOC→DOCX converter for the running platform: Word via COM or via AppleScript."""
    if platform == "win32":
        return WindowsWordConverter()
    if platform == "darwin":
        return MacWordConverter()
    return UnsupportedConverter()


class TemplatePreparer:
    def prepare(
        self,
        source_directory: Path,
        destination_directory: Path,
        *,
        converter: DocumentConverter | None = None,
    ) -> Path:
        source_directory = Path(source_directory)
        destination_directory = Path(destination_directory)
        if not source_directory.is_dir():
            raise TemplatePreparationError(f"Папка шаблонов не найдена: {source_directory}")

        sources = {
            "Талон": self._find_source(source_directory, "талон"),
            "Форма 4": self._find_source(source_directory, "форма 4"),
            "Форма 11": self._find_source(source_directory, "форма 11"),
        }
        missing = [name for name, path in sources.items() if path is None]
        if missing:
            raise TemplatePreparationError("Не найдены исходные шаблоны: " + ", ".join(missing))

        resolved_sources = {name: path for name, path in sources.items() if path is not None}
        fingerprints = {
            name: {"path": str(path.resolve()), "sha256": _sha256(path)}
            for name, path in resolved_sources.items()
        }
        manifest_path = destination_directory / "templates.json"
        expected_outputs = [destination_directory / f"{name}.docx" for name in resolved_sources]
        if manifest_path.is_file() and all(path.is_file() for path in expected_outputs):
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                manifest = {}
            if manifest.get("sources") == fingerprints:
                return destination_directory

        selected_converter = converter or default_converter()
        destination_directory.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(
            prefix=".timdoc-templates-", dir=destination_directory.parent
        ) as temp:
            work = Path(temp)
            transforms = {
                "Талон": _prepare_ticket_xml,
                "Форма 4": _prepare_form_xml,
                "Форма 11": _prepare_form_xml,
            }
            for name, transform in transforms.items():
                source = resolved_sources[name]
                converted = work / f"{name}-converted.docx"
                if source.suffix.casefold() == ".docx":
                    shutil.copyfile(source, converted)
                else:
                    selected_converter.convert(source, converted)
                _patch_docx(converted, work / f"{name}.docx", transform)

            manifest = {"schema_version": 1, "sources": fingerprints}
            (work / "templates.json").write_text(
                json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            destination_directory.mkdir(parents=True, exist_ok=True)
            for name in ("Талон.docx", "Форма 4.docx", "Форма 11.docx", "templates.json"):
                os.replace(work / name, destination_directory / name)
        return destination_directory

    @staticmethod
    def _find_source(
        directory: Path, phrase: str, *, suffixes: set[str] | None = None
    ) -> Path | None:
        allowed = suffixes or {".doc", ".docx"}
        candidates = [
            path
            for path in directory.iterdir()
            if path.is_file()
            and path.suffix.casefold() in allowed
            and phrase in _normalized(path.stem)
            and "мтк" not in _normalized(path.stem)
        ]
        return (
            sorted(candidates, key=lambda path: (_normalized(path.name), path.stat().st_size))[0]
            if candidates
            else None
        )


class WindowsWordConverter:
    def convert(self, source: Path, destination: Path) -> None:
        if sys.platform != "win32":
            raise TemplatePreparationError(
                "Преобразование DOC требует Windows и установленный Microsoft Word"
            )
        try:
            import win32com.client  # type: ignore[import-not-found, import-untyped]
        except ImportError as error:  # pragma: no cover - Windows packaging guard
            raise TemplatePreparationError(
                "В сборке отсутствует модуль автоматизации Word"
            ) from error

        source = source.resolve()
        destination = destination.resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        document = None
        try:
            document = word.Documents.Open(
                str(source), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
            )
            document.SaveAs2(str(destination), FileFormat=16, AddToRecentFiles=False)
        except Exception as error:
            raise TemplatePreparationError(
                f"Word не смог преобразовать {source.name}: {error}"
            ) from error
        finally:
            try:
                if document is not None:
                    document.Close(SaveChanges=False)
            except Exception:
                # Word has already written the file; a failing Close must not keep Word alive.
                pass
            finally:
                word.Quit()


# Word for Mac runs in the App Sandbox: every path outside its own containers triggers the
# "Grant File Access" dialog, once per file and per Word launch. Inside the Office group
# container Word reads and writes freely, so the source is staged there, converted in place
# and the result is moved out. No dialogs, nothing left behind.
OFFICE_GROUP_CONTAINER = Path("~/Library/Group Containers/UBF8T346G9.Office")
MAC_WORD_SCRIPT = """\
on run argv
  tell application "Microsoft Word"
    set display alerts to alerts none
    open file name (item 1 of argv)
    -- Right after opening Word may still be converting the file and rejects "save as".
    repeat with attempt from 1 to 15
      try
        save as active document file name (item 2 of argv) file format format document
        exit repeat
      on error message number code
        if attempt = 15 then error message number code
        delay 1
      end try
    end repeat
    close active document saving no
  end tell
end run
"""


class MacWordConverter:
    """Converts DOC to DOCX through Microsoft Word for Mac (development and demo use)."""

    def __init__(
        self,
        runner: Callable[..., Any] = subprocess.run,
        timeout: float = 120,
        staging_root: Path | None = None,
    ) -> None:
        self._runner = runner
        self._timeout = timeout
        self._staging_root = staging_root

    def convert(self, source: Path, destination: Path) -> None:
        if sys.platform != "darwin":
            raise TemplatePreparationError(
                "Преобразование DOC через AppleScript доступно только на macOS"
            )
        source = source.resolve()
        destination = destination.resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        root = self._staging_root or OFFICE_GROUP_CONTAINER.expanduser()
        if not root.is_dir():
            raise TemplatePreparationError(
                "Не найден Microsoft Word: отсутствует папка " + str(root)
            )
        staging_root = root / "Timdoc"
        try:
            staging_root.mkdir(exist_ok=True)
            staging = tempfile.TemporaryDirectory(prefix="convert-", dir=staging_root)
        except OSError as error:
            raise TemplatePreparationError(
                f"Нет доступа к папке Word для преобразования DOC: {error}"
            ) from error
        with staging as temp:
            staged_source = Path(temp) / source.name
            staged_result = Path(temp) / destination.name
            shutil.copyfile(source, staged_source)
            try:
                completed = self._runner(
                    ["osascript", "-e", MAC_WORD_SCRIPT, str(staged_source), str(staged_result)],
                    capture_output=True,
                    text=True,
                    timeout=self._timeout,
                    check=False,
                )
            except (OSError, subprocess.SubprocessError) as error:
                raise TemplatePreparationError(
                    f"Word не смог преобразовать {source.name}: {error}"
                ) from error
            if completed.returncode != 0 or not staged_result.is_file():
                detail = (completed.stderr or completed.stdout or "").strip() or "Word не ответил"
                raise TemplatePreparationError(
                    f"Word не смог преобразовать {source.name}: {detail}"
                )
            shutil.move(str(staged_result), destination)


class UnsupportedConverter:
    def convert(self, source: Path, destination: Path) -> None:
        raise TemplatePreparationError(
            "Преобразование DOC требует Windows или macOS с установленным Microsoft Word"
        )


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("w", W)

# Child order of w:rPr from the OOXML schema (CT_RPr); Word rejects out-of-order properties.
RPR_ORDER = (
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
)
BLANK_LINE = re.compile(r"_{3,}")
TICKET_SIGNATURE_LINE = 21
BODY_FONT_SIZE = "22"


Transform = Callable[[ET._Element, "StyleSheet"], None]


def _patch_docx(source: Path, destination: Path, transform: Transform) -> None:
    try:
        with zipfile.ZipFile(source) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
            styles = StyleSheet.from_archive(archive)
            transform(root, styles)
            _validate_namespace_hints(root)
            document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            with zipfile.ZipFile(destination, "w") as output:
                for info in archive.infolist():
                    data = (
                        document_xml if info.filename == "word/document.xml" else archive.read(info)
                    )
                    output.writestr(info, data)
    except (KeyError, ET.XMLSyntaxError, zipfile.BadZipFile, OSError) as error:
        raise TemplatePreparationError(f"Не удалось подготовить {source.name}: {error}") from error


class StyleSheet:
    """Enough of styles.xml to tell whether a run renders bold: the weight of the blank
    decides how wide its underscores are, and it usually comes from the paragraph style."""

    def __init__(self, bold_by_style: dict[str, bool | None], based_on: dict[str, str]) -> None:
        self._bold = bold_by_style
        self._based_on = based_on

    @classmethod
    def from_archive(cls, archive: zipfile.ZipFile) -> StyleSheet:
        try:
            root = ET.fromstring(archive.read("word/styles.xml"))
        except KeyError:
            return cls({}, {})
        bold: dict[str, bool | None] = {}
        based_on: dict[str, str] = {}
        for style in root.findall("./w:style", NS):
            style_id = style.get(f"{{{W}}}styleId")
            if not style_id:
                continue
            bold[style_id] = _bold_flag(style.find("./w:rPr", NS))
            parent = style.find("./w:basedOn", NS)
            if parent is not None and parent.get(f"{{{W}}}val"):
                based_on[style_id] = parent.get(f"{{{W}}}val")
        return cls(bold, based_on)

    def is_bold(self, paragraph: ET._Element, properties: ET._Element | None) -> bool:
        direct = _bold_flag(properties)
        if direct is not None:
            return direct
        style = paragraph.find("./w:pPr/w:pStyle", NS)
        style_id = style.get(f"{{{W}}}val") if style is not None else None
        seen: set[str] = set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            flag = self._bold.get(style_id)
            if flag is not None:
                return flag
            style_id = self._based_on.get(style_id)
        return False


def _bold_flag(properties: ET._Element | None) -> bool | None:
    if properties is None:
        return None
    bold = properties.find("./w:b", NS)
    if bold is None:
        return None
    return bold.get(f"{{{W}}}val", "1") not in ("0", "false")


def _prepare_ticket_xml(root: ET._Element, styles: StyleSheet) -> None:
    outer = root.find(".//w:body/w:tbl", NS)
    if outer is None:
        raise TemplatePreparationError("В талоне не найдена основная таблица")
    outer_cells = outer.findall("./w:tr/w:tc", NS)
    if len(outer_cells) != 2:
        raise TemplatePreparationError("Талон должен содержать две печатные копии")
    _equalize_copies(outer, outer_cells)
    for outer_cell in outer_cells:
        nested = outer_cell.find("./w:tbl", NS)
        if nested is None:
            raise TemplatePreparationError("В талоне не найдена таблица реквизитов")
        rows = nested.findall("./w:tr", NS)
        if len(rows) < 30:
            raise TemplatePreparationError("Структура таблицы талона не поддерживается")

        _set_ticket_value(rows[2], "{{ equipment_name }}")
        serial_cells = rows[4].findall("./w:tc", NS)[1:-1]
        if len(serial_cells) < 14:
            raise TemplatePreparationError("В талоне меньше 14 ячеек заводского номера")
        for index, cell in enumerate(serial_cells[:14]):
            _set_element_text(cell, f"{{{{ serial_{index:02d} }}}}")
        _set_ticket_value(rows[9], "{{ customer_name }}")
        _set_ticket_value(rows[11], "{{ customer_inn }}")
        _set_ticket_value(rows[13], "{{ customer_representative }}")
        _set_ticket_pair(rows[15], "{{ customer_country }}", "{{ customer_postal_code }}")
        _set_ticket_pair(rows[17], "{{ customer_region }}", "{{ customer_district }}")
        _set_ticket_value(rows[19], "{{ customer_locality }}")
        _set_ticket_pair(rows[21], "{{ customer_street }}", "{{ customer_house }}")
        _set_ticket_value(rows[23], "{{ customer_email }}")
        _set_ticket_pair(rows[25], "", "{{ customer_phone }}")
        _set_ticket_value(rows[27], "{{ service_center }}")
        # Строка 29 «Дата постановки на гарантийный учет» остаётся пустой: её заполняют от руки.

        direct_paragraphs = outer_cell.findall("./w:p", NS)
        service_index = next(
            (
                index
                for index, paragraph in enumerate(direct_paragraphs)
                if "СЕРВИСНЫЙ ЦЕНТР" in _element_text(paragraph)
            ),
            -1,
        )
        if service_index >= 0:
            signature = next(
                (
                    paragraph
                    for paragraph in direct_paragraphs[service_index + 1 :]
                    if "_" in _element_text(paragraph)
                ),
                None,
            )
            if signature is not None and not _fill_blank(
                signature,
                lambda _blank, _properties: [(" {{ service_employee }} ", _plain_style)],
                keep_left=TICKET_SIGNATURE_LINE,
            ):
                _set_element_text(signature, "_____________________ {{ service_employee }} М.П.")
        # Фамилия набрана размером основного текста, как в заполненном образце; чтобы талон
        # остался на одной странице, хвостовые пустые абзацы ячейки убираются так же, как в нём.
        _drop_trailing_empty_paragraphs(outer_cell)


def _equalize_copies(outer: ET._Element, outer_cells: list[ET._Element]) -> None:
    """Give both printed copies the same width: the source template has a narrower right cell,
    so long values wrapped only in one copy and the halves ended up different heights."""
    columns = outer.findall("./w:tblGrid/w:gridCol", NS)
    widths = [int(column.get(f"{{{W}}}w") or 0) for column in columns]
    if len(widths) != 2 or not all(widths):
        return
    total = sum(widths)
    shares = (total // 2, total - total // 2)
    for column, share in zip(columns, shares, strict=True):
        column.set(f"{{{W}}}w", str(share))
    for cell, share in zip(outer_cells, shares, strict=True):
        width = cell.find("./w:tcPr/w:tcW", NS)
        if width is not None and width.get(f"{{{W}}}type", "dxa") == "dxa":
            width.set(f"{{{W}}}w", str(share))


def _drop_trailing_empty_paragraphs(cell: ET._Element) -> None:
    children = list(cell)
    while len(children) >= 2:
        last, previous = children[-1], children[-2]
        if ET.QName(last).localname != "p" or ET.QName(previous).localname != "p":
            break
        if _element_text(last).strip() or last.find(".//w:drawing", NS) is not None:
            break
        cell.remove(last)
        children.pop()


# Каждое поле формы: подпись под линией -> имя переменной. Линия над подписью заполняется
# значением, отцентрованным в подчёркиваниях той же ширины (фильтр blank в генераторе).
# «Дата составления» и контактное лицо потребителя намеренно не заполняются: в подписанных
# образцах их вписывают от руки при подписании.
FORM_FIELDS = {
    "место составления": "service_location",
    "(наименование организации)": "service_center",
    "(марка продукции)": "equipment_name",
    "(номер)": "serial_number",
    "(наименование организации, инн)": "customer_name_inn",
    "(адрес организации)": "customer_address",
}
REQUIRED_FORM_FIELDS = {
    "(наименование организации)",
    "(марка продукции)",
    "(номер)",
    "(наименование организации, инн)",
    "(адрес организации)",
}


def _prepare_form_xml(root: ET._Element, styles: StyleSheet) -> None:
    paragraphs = root.findall(".//w:p", NS)
    found: set[str] = set()
    for index, paragraph in enumerate(paragraphs):
        label = _normalized(_element_text(paragraph))
        variable = FORM_FIELDS.get(label)
        if variable is None:
            continue
        target = _previous_nonempty(paragraphs, index)
        if target is None:
            raise TemplatePreparationError(f"Нет поля перед подписью {label}")
        if not _fill_blank(target, _line_placeholder(variable, target, styles)):
            raise TemplatePreparationError(f"Над подписью {label} нет линии для заполнения")
        found.add(label)
    missing = REQUIRED_FORM_FIELDS - found
    if missing:
        raise TemplatePreparationError("В форме не найдены поля: " + ", ".join(sorted(missing)))

    for index, paragraph in enumerate(paragraphs):
        label = _normalized(_element_text(paragraph))
        if "(должность)" in label:
            target = _previous_nonempty(paragraphs, index)
            if target is not None:
                _fill_blank(target, _line_placeholder("service_employee_position", target, styles))
        if "м.п." in label and "(ф.и.о.)" in label:
            target = _previous_nonempty(paragraphs, index)
            if target is not None:
                # Первый блок подчёркиваний отведён под подпись, второй под фамилию исполнителя.
                _fill_blank(
                    target, _line_placeholder("service_employee", target, styles), occurrence=1
                )


Segment = tuple[str, "Style | None"]
Style = Callable[[ET._Element | None], ET._Element | None]


def _line_placeholder(
    variable: str, paragraph: ET._Element, styles: StyleSheet
) -> Callable[[str, ET._Element | None], list[Segment]]:
    """Three runs: underscores in the blank's own style on both sides, the value in body style.

    The padding keeps the blank's size and weight (a bold style makes underscores wider),
    so the generator can compute how many of them the value displaces (filters blank_left
    and blank_right) and the line keeps its printed length.
    """

    def placeholder(blank: str, properties: ET._Element | None) -> list[Segment]:
        size = _font_size(properties)
        bold = "True" if styles.is_bold(paragraph, properties) else "False"
        spec = f"{len(blank)}, {size}, {bold}"
        return [
            (f"{{{{ {variable} | blank_left({spec}) }}}}", None),
            (f"{{{{ {variable} }}}}", _value_style),
            (f"{{{{ {variable} | blank_right({spec}) }}}}", None),
        ]

    return placeholder


def _font_size(properties: ET._Element | None) -> int:
    size = properties.find("./w:sz", NS) if properties is not None else None
    try:
        return int(size.get(f"{{{W}}}val")) if size is not None else int(BODY_FONT_SIZE)
    except (TypeError, ValueError):
        return int(BODY_FONT_SIZE)


def _fill_blank(
    paragraph: ET._Element,
    replacement: Callable[[str, ET._Element | None], list[Segment]],
    *,
    pattern: re.Pattern[str] = BLANK_LINE,
    occurrence: int = 0,
    keep_left: int = 0,
) -> bool:
    """Replace one blank (a run of underscores) inside a paragraph, keeping everything else.

    The paragraph is flattened into tokens (characters, tabs, opaque elements) that remember
    their run properties, the matched blank is swapped for the replacement segments (each with
    its own style, None = the blank's own properties), and the runs are rebuilt so label
    prefixes, tab stops and the neighbouring blanks survive untouched.
    Returns False when the paragraph has no such blank.
    """
    runs = paragraph.findall("./w:r", NS)
    if not runs:
        return False
    tokens: list[tuple[str | ET._Element, ET._Element | None]] = []
    for run in runs:
        properties = run.find("./w:rPr", NS)
        for child in run:
            tag = ET.QName(child).localname
            if tag == "rPr":
                continue
            if tag == "t":
                tokens.extend((char, properties) for char in child.text or "")
            elif tag == "tab":
                tokens.append(("\t", properties))
            elif tag != "lastRenderedPageBreak":
                tokens.append((child, properties))
    text = "".join(token if isinstance(token, str) else "\x00" for token, _ in tokens)
    matches = list(pattern.finditer(text))
    if occurrence >= len(matches):
        return False
    match = matches[occurrence]
    start = match.start() + keep_left
    end = match.end()
    if start >= end:
        return False
    original_properties = tokens[start][1]
    replaced: list[tuple[str | ET._Element, ET._Element | None]] = []
    for segment_text, style in replacement(text[start:end], original_properties):
        segment_properties = (style or copy.deepcopy)(original_properties)
        replaced.extend((char, segment_properties) for char in segment_text)
    tokens = tokens[:start] + replaced + tokens[end:]

    insert_at = list(paragraph).index(runs[0])
    for run in runs:
        paragraph.remove(run)
    for run in _build_runs(tokens):
        paragraph.insert(insert_at, run)
        insert_at += 1
    return True


def _build_runs(
    tokens: list[tuple[str | ET._Element, ET._Element | None]],
) -> list[ET._Element]:
    runs: list[ET._Element] = []
    current_properties: ET._Element | None = None
    run: ET._Element | None = None
    text_element: ET._Element | None = None
    for token, properties in tokens:
        if run is None or properties is not current_properties:
            run = ET.Element(f"{{{W}}}r")
            if properties is not None:
                run.append(copy.deepcopy(properties))
            runs.append(run)
            current_properties = properties
            text_element = None
        if isinstance(token, str) and token == "\t":
            ET.SubElement(run, f"{{{W}}}tab")
            text_element = None
        elif isinstance(token, str):
            if text_element is None:
                text_element = ET.SubElement(run, f"{{{W}}}t")
                text_element.set(XML_SPACE, "preserve")
                text_element.text = ""
            text_element.text = (text_element.text or "") + token
        else:
            run.append(token)
            text_element = None
    return runs


def _value_style(properties: ET._Element | None) -> ET._Element:
    """Run properties for a filled-in value: body size, regular weight, underlined line."""
    result = copy.deepcopy(properties) if properties is not None else ET.Element(f"{{{W}}}rPr")
    # Paragraph styles of some labels ("Место составления") are bold; values never are.
    _set_run_property(result, "b", val="0")
    _set_run_property(result, "bCs", val="0")
    _set_run_property(result, "sz", val=BODY_FONT_SIZE)
    _set_run_property(result, "szCs", val=BODY_FONT_SIZE)
    _set_run_property(result, "u", val="single")
    return result


def _plain_style(properties: ET._Element | None) -> ET._Element | None:
    """Run properties for the ticket signature: style size, regular weight, on the line."""
    result = copy.deepcopy(properties) if properties is not None else ET.Element(f"{{{W}}}rPr")
    for tag in ("b", "bCs", "sz", "szCs"):
        for element in result.findall(f"./w:{tag}", NS):
            result.remove(element)
    _set_run_property(result, "u", val="single")
    return result


def _set_run_property(properties: ET._Element, tag: str, **attributes: str) -> None:
    for element in properties.findall(f"./w:{tag}", NS):
        properties.remove(element)
    element = ET.Element(f"{{{W}}}{tag}")
    for name, value in attributes.items():
        element.set(f"{{{W}}}{name}", value)
    rank = RPR_ORDER.index(tag)
    position = len(properties)
    for index, sibling in enumerate(properties):
        sibling_tag = ET.QName(sibling).localname
        if sibling_tag in RPR_ORDER and RPR_ORDER.index(sibling_tag) > rank:
            position = index
            break
    properties.insert(position, element)


def _set_ticket_value(row: ET._Element, value: str) -> None:
    cells = row.findall("./w:tc", NS)
    if len(cells) < 3:
        raise TemplatePreparationError("В строке талона нет ячейки для значения")
    _set_element_text(cells[1], value)


def _set_ticket_pair(row: ET._Element, first: str, second: str) -> None:
    cells = row.findall("./w:tc", NS)
    if len(cells) < 4:
        raise TemplatePreparationError("В строке талона нет двух ячеек для значений")
    _set_element_text(cells[1], first)
    _set_element_text(cells[-2], second)


def _set_element_text(element: ET._Element, value: str) -> None:
    texts = element.findall(".//w:t", NS)
    if texts:
        texts[0].text = value
        for text in texts[1:]:
            text.text = ""
        return
    paragraph = element.find(".//w:p", NS)
    if paragraph is None:
        paragraph = ET.SubElement(element, f"{{{W}}}p")
    run = ET.SubElement(paragraph, f"{{{W}}}r")
    # An empty cell keeps its font only on the paragraph mark (pPr/rPr); a new run
    # would otherwise fall back to the document default (Calibri instead of Tahoma).
    mark_properties = paragraph.find("./w:pPr/w:rPr", NS)
    if mark_properties is not None:
        run.append(copy.deepcopy(mark_properties))
    text = ET.SubElement(run, f"{{{W}}}t")
    text.set(XML_SPACE, "preserve")
    text.text = value


def _element_text(element: ET._Element) -> str:
    return "".join(text.text or "" for text in element.findall(".//w:t", NS))


def _previous_nonempty(paragraphs: list[ET._Element], index: int) -> ET._Element | None:
    for candidate in reversed(paragraphs[:index]):
        if _element_text(candidate).strip():
            return candidate
    return None


def _normalized(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).casefold().replace("ё", "е")
    return " ".join(value.replace("\xa0", " ").split())


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validate_namespace_hints(root: ET._Element) -> None:
    compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    required = set((root.get(f"{{{compatibility}}}Ignorable") or "").split())
    declared = {prefix for prefix in root.nsmap if prefix}
    missing = required - declared
    if missing:
        raise TemplatePreparationError(
            "В DOCX потеряны служебные пространства имен: " + ", ".join(sorted(missing))
        )
