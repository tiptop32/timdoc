from pathlib import Path

import pytest

from timdoc_act_parser import ActParseError, ActParser

REAL_ACT = Path("/Users/tiptop32/my_git_reps/timdoc/doc_templates/Акт.pdf")


@pytest.mark.skipif(not REAL_ACT.exists(), reason="real acceptance fixture is outside Git")
def test_parses_real_saby_act_into_six_equipment_records() -> None:
    act = ActParser().parse(REAL_ACT)

    assert act.customer_name == "ООО «Рощинский»"
    assert act.service_center_name == "ООО «Акрос РБ»"
    assert act.document_date == "2026-09-07"
    assert [item.serial_number for item in act.equipment] == [
        "M0S07013002106",
        "M0S07013002105",
        "M0S07013002107",
        "M0S07013002108",
        "M0S07013002109",
        "M0S07013002110",
    ]
    assert {item.name for item in act.equipment} == {"RSM SS-780-13"}
    assert act.customer_inn == ""
    assert any("ИНН" in warning for warning in act.warnings)


def test_rejects_non_pdf(tmp_path: Path) -> None:
    source = tmp_path / "act.txt"
    source.write_text("not a pdf", encoding="utf-8")

    with pytest.raises(ActParseError, match="PDF"):
        ActParser().parse(source)


SYNTHETIC_ACT = """
Акт приема-передачи товара
г. Уфа «07» сентября 2026 г.
Общество с ограниченной ответственностью (ООО «Акрос РБ»), именуемое в дальнейшем «Продавец»,
и Общество (ООО «Рощинский»), именуемое в дальнейшем «Покупатель», ИНН 0268104130,
составили настоящий акт о том, что Продавец передал следующий Товар:
№ Наименование Заводской номер Год Изготовитель
1 Сеялка RSM SS-780-13 M0S07013002106 2026 АО «Клевер»
2 Сеялка RSM SS-780-13 M0S07013002105 2026 АО «Клевер»
3 Плуг ПЛН-5-35 PLN535ABC0001 2025 ООО «Завод»
4 Сеялка RSM SS-780-13 M0S07013002106 2026 АО «Клевер»
2. Товар передан в исправном состоянии.
"""


class FakePage:
    def __init__(self, text: str, *, layout: bool = True) -> None:
        self.text = text
        self.layout = layout

    def extract_text(self, **options: object) -> str:
        if options and not self.layout:
            raise TypeError("extraction_mode is not supported")
        return self.text


class FakeReader:
    def __init__(self, pages: list[FakePage], *, encrypted: bool = False) -> None:
        self.pages = pages
        self.is_encrypted = encrypted

    def decrypt(self, password: str) -> int:
        return 0


def _reader(monkeypatch: pytest.MonkeyPatch, reader: FakeReader) -> None:
    from timdoc_act_parser import parser as module

    monkeypatch.setattr(module, "PdfReader", lambda path: reader)


def test_missing_pdf_is_reported(tmp_path: Path) -> None:
    with pytest.raises(ActParseError, match="не найден"):
        ActParser().parse(tmp_path / "missing.pdf")


def test_encrypted_and_scanned_pdfs_are_rejected(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "act.pdf"
    source.write_bytes(b"%PDF-1.4")

    _reader(monkeypatch, FakeReader([], encrypted=True))
    with pytest.raises(ActParseError, match="паролем"):
        ActParser().parse(source)

    _reader(monkeypatch, FakeReader([FakePage("   ")]))
    with pytest.raises(ActParseError, match="OCR"):
        ActParser().parse(source)


def test_synthetic_act_is_parsed_with_fallbacks(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "act.pdf"
    source.write_bytes(b"%PDF-1.4")
    _reader(monkeypatch, FakeReader([FakePage(SYNTHETIC_ACT, layout=False)]))

    act = ActParser().parse(source)

    assert act.document_date == "2026-09-07"
    assert act.customer_name == "ООО «Рощинский»"
    assert act.service_center_name == "ООО «Акрос РБ»"
    assert act.customer_inn == "0268104130"
    assert [item.serial_number for item in act.equipment] == [
        "M0S07013002106",
        "M0S07013002105",
        "PLN535ABC0001",
    ]
    assert act.equipment[0].name == "RSM SS-780-13"
    assert act.equipment[0].manufacture_year == "2026"
    assert act.equipment[0].manufacturer == "АО «Клевер»"
    assert "ПЛН-5-35" in act.equipment[2].name
    assert act.warnings == []


def test_act_without_recognizable_parties_reports_warnings(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "act.pdf"
    source.write_bytes(b"%PDF-1.4")
    text = "Документ без даты и сторон. " * 4 + "\nСтроки без техники и без номеров.\n"
    _reader(monkeypatch, FakeReader([FakePage(text)]))

    act = ActParser().parse(source)

    assert act.document_date == ""
    assert act.customer_name == ""
    assert act.equipment == []
    assert len(act.warnings) == 5


def test_reader_failures_become_parse_errors(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "act.pdf"
    source.write_bytes(b"not really a pdf")
    from timdoc_act_parser import parser as module

    def broken(path: Path) -> FakeReader:
        raise module.PdfReadError("EOF marker not found")

    monkeypatch.setattr(module, "PdfReader", broken)
    with pytest.raises(ActParseError, match="Не удалось прочитать PDF"):
        ActParser().parse(source)
