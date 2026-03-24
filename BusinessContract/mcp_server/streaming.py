"""Streaming Excel cell reader using lxml iterparse + shared strings."""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from typing import Iterator

from lxml import etree

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


@dataclass
class CellInfo:
    """One cell's data."""
    ref: str          # e.g. "B3"
    row: int
    col: int
    value: str | None = None
    formula: str | None = None
    cell_type: str = "n"  # n=number, s=shared-string, str=inline, b=bool


_COL_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$")


def col_letter_to_index(letters: str) -> int:
    """A->1, B->2, ..., Z->26, AA->27."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def index_to_col_letter(idx: int) -> str:
    """1->A, 2->B, ..., 27->AA."""
    result = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result.append(chr(rem + ord("A")))
    return "".join(reversed(result))


def parse_ref(ref: str) -> tuple[int, int]:
    """Parse 'B3' or '$B$3' -> (row=3, col=2). Strips $ signs."""
    m = _COL_RE.match(ref)
    if not m:
        return 0, 0
    return int(m.group(2)), col_letter_to_index(m.group(1))


def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Load shared string table from xl/sharedStrings.xml."""
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    strings: list[str] = []
    for event, elem in etree.iterparse(
        io.BytesIO(data), events=("end",),
        tag=f"{{{_NS}}}si",
    ):
        # Concatenate all <t> text within this <si>
        parts = []
        for t_elem in elem.iter(f"{{{_NS}}}t"):
            if t_elem.text:
                parts.append(t_elem.text)
        strings.append("".join(parts))
        elem.clear()
    return strings


def _get_sheet_xml_path(zf: zipfile.ZipFile, sheet_name: str) -> str | None:
    """Map sheet name -> xl/worksheets/sheet*.xml path."""
    try:
        wb_xml = zf.read("xl/workbook.xml")
    except KeyError:
        return None

    # Parse sheet name -> rId mapping
    rid_map: dict[str, str] = {}
    for _, elem in etree.iterparse(
        io.BytesIO(wb_xml), events=("end",),
    ):
        local = etree.QName(elem.tag).localname
        if local == "sheet":
            name = elem.get("name", "")
            rid = elem.get(f"{{{_NS.replace('spreadsheetml', 'officeDocument/2006/relationships')}}}id")
            if not rid:
                rid = elem.get(r"{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if name == sheet_name and rid:
                rid_map[rid] = name
            elem.clear()

    if not rid_map:
        return None

    # Parse workbook rels to get sheet path
    try:
        rels_xml = zf.read("xl/_rels/workbook.xml.rels")
    except KeyError:
        return None

    for _, elem in etree.iterparse(io.BytesIO(rels_xml), events=("end",)):
        local = etree.QName(elem.tag).localname
        if local == "Relationship":
            rid = elem.get("Id", "")
            if rid in rid_map:
                target = elem.get("Target", "")
                if not target.startswith("/"):
                    target = "xl/" + target
                else:
                    target = target.lstrip("/")
                return target
            elem.clear()

    return None


def stream_sheet_cells(
    zf: zipfile.ZipFile,
    sheet_name: str,
    shared_strings: list[str] | None = None,
) -> Iterator[CellInfo]:
    """Stream cells from a sheet, resolving shared strings."""
    if shared_strings is None:
        shared_strings = _load_shared_strings(zf)

    xml_path = _get_sheet_xml_path(zf, sheet_name)
    if not xml_path:
        return

    try:
        data = zf.read(xml_path)
    except KeyError:
        return

    for event, elem in etree.iterparse(
        io.BytesIO(data), events=("end",),
        tag=f"{{{_NS}}}c",
    ):
        ref = elem.get("r", "")
        cell_type = elem.get("t", "n")
        row, col = parse_ref(ref)

        # Value
        v_elem = elem.find(f"{{{_NS}}}v")
        value = v_elem.text if v_elem is not None and v_elem.text else None

        # Resolve shared string
        if cell_type == "s" and value is not None:
            try:
                value = shared_strings[int(value)]
            except (ValueError, IndexError):
                pass

        # Formula
        f_elem = elem.find(f"{{{_NS}}}f")
        formula = f_elem.text if f_elem is not None and f_elem.text else None

        yield CellInfo(
            ref=ref, row=row, col=col,
            value=value, formula=formula, cell_type=cell_type,
        )
        elem.clear()


def read_cell_neighborhood(
    zf: zipfile.ZipFile,
    sheet_name: str,
    center_ref: str,
    radius: int = 3,
    shared_strings: list[str] | None = None,
) -> dict[str, CellInfo]:
    """Read cells within ±radius rows/cols of center_ref.

    Returns dict keyed by cell ref (e.g. "B3").
    """
    crow, ccol = parse_ref(center_ref)
    if crow == 0:
        return {}

    min_row = max(1, crow - radius)
    max_row = crow + radius
    min_col = max(1, ccol - radius)
    max_col = ccol + radius

    result: dict[str, CellInfo] = {}
    for cell in stream_sheet_cells(zf, sheet_name, shared_strings):
        if min_row <= cell.row <= max_row and min_col <= cell.col <= max_col:
            result[cell.ref] = cell
    return result


def get_sheet_summary(
    zf: zipfile.ZipFile,
    sheet_name: str,
    shared_strings: list[str] | None = None,
    max_rows: int = 5,
) -> dict:
    """Get sheet summary: dimensions, header row, sample rows.

    Streams cells but only keeps first (1 + max_rows) rows in memory.
    Tracks max_row/max_col without storing every cell.
    """
    # Only store the rows we need (header + sample rows)
    keep_rows: dict[tuple[int, int], CellInfo] = {}
    max_keep = 1 + max_rows  # row 1 (header) + max_rows data rows
    max_r = max_c = 0

    for cell in stream_sheet_cells(zf, sheet_name, shared_strings):
        max_r = max(max_r, cell.row)
        max_c = max(max_c, cell.col)
        if cell.row <= max_keep:
            keep_rows[(cell.row, cell.col)] = cell

    # Extract header (row 1)
    headers = []
    for c in range(1, max_c + 1):
        ci = keep_rows.get((1, c))
        headers.append(ci.value if ci else None)

    # Sample rows
    sample_rows = []
    for r in range(2, min(2 + max_rows, max_r + 1)):
        row_data = []
        for c in range(1, max_c + 1):
            ci = keep_rows.get((r, c))
            row_data.append(ci.value if ci else None)
        sample_rows.append(row_data)

    return {
        "sheet_name": sheet_name,
        "max_row": max_r,
        "max_col": max_c,
        "headers": headers,
        "sample_rows": sample_rows,
    }


def list_sheets(zf: zipfile.ZipFile) -> list[str]:
    """List all sheet names from workbook.xml."""
    try:
        wb_xml = zf.read("xl/workbook.xml")
    except KeyError:
        return []

    sheets: list[str] = []
    for _, elem in etree.iterparse(
        io.BytesIO(wb_xml), events=("end",),
    ):
        local = etree.QName(elem.tag).localname
        if local == "sheet":
            name = elem.get("name")
            if name:
                sheets.append(name)
            elem.clear()
    return sheets


def get_named_ranges(zf: zipfile.ZipFile) -> list[dict]:
    """Get all defined names from workbook.xml."""
    try:
        wb_xml = zf.read("xl/workbook.xml")
    except KeyError:
        return []

    names: list[dict] = []
    for _, elem in etree.iterparse(
        io.BytesIO(wb_xml), events=("end",),
    ):
        local = etree.QName(elem.tag).localname
        if local == "definedName":
            names.append({
                "name": elem.get("name", ""),
                "value": (elem.text or "").strip(),
                "scope_sheet_id": elem.get("localSheetId"),
                "hidden": elem.get("hidden", "0") == "1",
            })
            elem.clear()
    return names
