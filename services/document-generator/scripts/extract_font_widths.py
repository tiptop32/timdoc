"""Извлекает ширины глифов из TrueType-шрифта в JSON для фильтра blank.

Использование:
    python scripts/extract_font_widths.py /System/Library/Fonts/Supplemental/Tahoma.ttf \\
        src/timdoc_document_generator/tahoma_widths.json

Читаются таблицы head (unitsPerEm), hhea/hmtx (advance widths) и cmap format 4.
Сохраняются только символы, которые встречаются в русских документах: ASCII, латиница-1,
кириллица и типографская пунктуация. Значения в единицах шрифта (для Tahoma 2048 на em).
"""

from __future__ import annotations

import json
import struct
import sys
from pathlib import Path

RANGES = ((0x20, 0x7E), (0xA0, 0x17F), (0x400, 0x45F), (0x490, 0x491), (0x2010, 0x2122))


def _tables(data: bytes) -> dict[str, tuple[int, int]]:
    count = struct.unpack(">H", data[4:6])[0]
    tables = {}
    for index in range(count):
        offset = 12 + index * 16
        tag = data[offset : offset + 4].decode("latin-1")
        start, length = struct.unpack(">II", data[offset + 8 : offset + 16])
        tables[tag] = (start, length)
    return tables


def _cmap_format4(data: bytes, start: int) -> dict[int, int]:
    count = struct.unpack(">H", data[start + 2 : start + 4])[0]
    for index in range(count):
        platform, encoding, offset = struct.unpack(
            ">HHI", data[start + 4 + index * 8 : start + 12 + index * 8]
        )
        sub = start + offset
        if struct.unpack(">H", data[sub : sub + 2])[0] != 4 or (platform, encoding) not in (
            (3, 1),
            (0, 3),
            (0, 4),
        ):
            continue
        segments = struct.unpack(">H", data[sub + 6 : sub + 8])[0] // 2
        ends = struct.unpack(f">{segments}H", data[sub + 14 : sub + 14 + segments * 2])
        starts_at = sub + 16 + segments * 2
        starts = struct.unpack(f">{segments}H", data[starts_at : starts_at + segments * 2])
        deltas_at = starts_at + segments * 2
        deltas = struct.unpack(f">{segments}h", data[deltas_at : deltas_at + segments * 2])
        ranges_at = deltas_at + segments * 2
        range_offsets = struct.unpack(f">{segments}H", data[ranges_at : ranges_at + segments * 2])
        mapping: dict[int, int] = {}
        for segment in range(segments):
            for code in range(starts[segment], min(ends[segment], 0xFFFE) + 1):
                if range_offsets[segment] == 0:
                    glyph = (code + deltas[segment]) & 0xFFFF
                else:
                    address = (
                        ranges_at
                        + segment * 2
                        + range_offsets[segment]
                        + (code - starts[segment]) * 2
                    )
                    glyph = struct.unpack(">H", data[address : address + 2])[0]
                    if glyph:
                        glyph = (glyph + deltas[segment]) & 0xFFFF
                if glyph:
                    mapping[code] = glyph
        return mapping
    raise SystemExit("В шрифте нет cmap format 4 для Unicode")


def extract(font_path: Path) -> dict[str, object]:
    data = font_path.read_bytes()
    tables = _tables(data)
    units_per_em = struct.unpack(">H", data[tables["head"][0] + 18 : tables["head"][0] + 20])[0]
    metrics_count = struct.unpack(">H", data[tables["hhea"][0] + 34 : tables["hhea"][0] + 36])[0]
    hmtx = tables["hmtx"][0]
    advances = [
        struct.unpack(">H", data[hmtx + index * 4 : hmtx + index * 4 + 2])[0]
        for index in range(metrics_count)
    ]
    cmap = _cmap_format4(data, tables["cmap"][0])
    widths: dict[str, int] = {}
    for low, high in RANGES:
        for code in range(low, high + 1):
            glyph = cmap.get(code)
            if glyph is None:
                continue
            widths[chr(code)] = advances[min(glyph, metrics_count - 1)]
    return {"font": font_path.name, "units_per_em": units_per_em, "widths": widths}


def main() -> None:
    source, destination = Path(sys.argv[1]), Path(sys.argv[2])
    payload = extract(source)
    destination.write_text(
        json.dumps(payload, ensure_ascii=False, indent=0, sort_keys=True) + "\n", encoding="utf-8"
    )
    print(f"{len(payload['widths'])} символов, {payload['units_per_em']} единиц на em")


if __name__ == "__main__":
    main()
