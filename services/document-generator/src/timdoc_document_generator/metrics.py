"""Ширины символов Tahoma для подгонки значений под линии подчёркиваний.

Таблицы `tahoma_widths.json` и `tahoma_bold_widths.json` извлечены из Tahoma.ttf и
Tahoma Bold.ttf скриптом `scripts/extract_font_widths.py` и не зависят от того, установлен
ли шрифт там, где работает приложение. Ширины в единицах шрифта (2048 на em) для одного
кегля; для другого кегля масштабируются линейно.
"""

from __future__ import annotations

import json
from functools import lru_cache
from importlib import resources

UNDERSCORE = "_"


@lru_cache(maxsize=2)
def _table(bold: bool) -> dict[str, int]:
    name = "tahoma_bold_widths.json" if bold else "tahoma_widths.json"
    payload = json.loads(resources.files(__package__).joinpath(name).read_text(encoding="utf-8"))
    widths: dict[str, int] = payload["widths"]
    return widths


def char_width(char: str, *, bold: bool = False) -> int:
    """Advance width of one character; unknown characters count as an average letter."""
    widths = _table(bold)
    return widths.get(char, widths["n"])


def text_width(text: str, *, bold: bool = False) -> int:
    return sum(char_width(char, bold=bold) for char in text)


def underscores_for(width: float, *, bold: bool = False) -> int:
    """How many whole underscores fit into the given width (never more: a line that grows
    even by a few points wraps and drags the tab-aligned blocks after it)."""
    return max(0, int(width // char_width(UNDERSCORE, bold=bold)))
