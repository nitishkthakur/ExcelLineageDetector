"""Fast streaming scanner for upstream tracing.

Two modes:
- Model scanning: only hardcoded (non-formula) numeric values
- Upstream scanning: ALL numeric values (formula + hardcoded)

Uses lxml iterparse for O(n) streaming — constant memory regardless of file size.
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from lxml import etree

from lineage.hardcoded_scanner import (
    _col_to_idx,
    _idx_to_col,
    _CELL_RE,
    _runs,
    _get_sheet_map,
)
from lineage.tracing.models import TracingVector

MIN_VECTOR_LEN = 3


# ---------------------------------------------------------------------------
# Streaming parsers
# ---------------------------------------------------------------------------

def _stream_all_numerics(data: bytes) -> list[tuple[int, int, float]]:
    """Stream-parse sheet XML and return (row, col, value) for ALL numeric cells.

    Captures both formula-derived and hardcoded values — needed for upstream
    files where an analyst may have copied a formula result.
    """
    results: list[tuple[int, int, float]] = []
    in_cell = False
    cell_ref = ""
    cell_type = "n"
    pending_value: float | None = None

    try:
        context = etree.iterparse(
            io.BytesIO(data),
            events=("start", "end"),
            recover=True,
            no_network=True,
        )
        for event, elem in context:
            ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

            if event == "start":
                if ltag == "c":
                    in_cell = True
                    pending_value = None
                    cell_ref = elem.get("r", "")
                    cell_type = elem.get("t", "n")
            else:  # "end"
                if ltag == "v" and in_cell:
                    if cell_type not in ("s", "b", "e", "str") and elem.text:
                        try:
                            pending_value = float(elem.text)
                        except ValueError:
                            pass
                elif ltag == "c":
                    if in_cell and pending_value is not None and cell_ref:
                        m = _CELL_RE.match(cell_ref)
                        if m:
                            results.append((
                                int(m.group(2)),
                                _col_to_idx(m.group(1)),
                                pending_value,
                            ))
                    in_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()

        del context
    except Exception:
        pass

    return results


def _stream_hardcoded_numerics(data: bytes) -> list[tuple[int, int, float]]:
    """Stream-parse sheet XML for hardcoded (non-formula) numeric cells only."""
    results: list[tuple[int, int, float]] = []
    in_cell = False
    has_formula = False
    cell_ref = ""
    cell_type = "n"
    pending_value: float | None = None

    try:
        context = etree.iterparse(
            io.BytesIO(data),
            events=("start", "end"),
            recover=True,
            no_network=True,
        )
        for event, elem in context:
            ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

            if event == "start":
                if ltag == "c":
                    in_cell = True
                    has_formula = False
                    pending_value = None
                    cell_ref = elem.get("r", "")
                    cell_type = elem.get("t", "n")
                elif ltag == "f" and in_cell:
                    has_formula = True
            else:  # "end"
                if ltag == "v" and in_cell and not has_formula:
                    if cell_type not in ("s", "b", "e", "str") and elem.text:
                        try:
                            pending_value = float(elem.text)
                        except ValueError:
                            pass
                elif ltag == "c":
                    if in_cell and not has_formula and pending_value is not None and cell_ref:
                        m = _CELL_RE.match(cell_ref)
                        if m:
                            results.append((
                                int(m.group(2)),
                                _col_to_idx(m.group(1)),
                                pending_value,
                            ))
                    in_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()

        del context
    except Exception:
        pass

    return results


# ---------------------------------------------------------------------------
# Vector construction
# ---------------------------------------------------------------------------

def _cells_to_vectors(
    cells: list[tuple[int, int, float]],
    file_name: str,
    sheet: str,
    min_len: int,
) -> list[TracingVector]:
    """Convert a cell list to TracingVectors with full value lists."""
    vectors: list[TracingVector] = []

    # --- Column vectors: same column, consecutive rows ---
    by_col: dict[int, list[tuple[int, float]]] = {}
    for row, col, val in cells:
        by_col.setdefault(col, []).append((row, val))

    for col, items in by_col.items():
        items.sort()
        col_letter = _idx_to_col(col)
        for start_r, end_r, vals in _runs(items, min_len):
            vectors.append(TracingVector(
                file=file_name,
                sheet=sheet,
                cell_range=f"{col_letter}{start_r}:{col_letter}{end_r}",
                direction="column",
                length=len(vals),
                start_cell=f"{col_letter}{start_r}",
                end_cell=f"{col_letter}{end_r}",
                values=tuple(vals),
            ))

    # --- Row vectors: same row, consecutive columns ---
    by_row: dict[int, list[tuple[int, float]]] = {}
    for row, col, val in cells:
        by_row.setdefault(row, []).append((col, val))

    for row, items in by_row.items():
        items.sort()
        for start_c, end_c, vals in _runs(items, min_len):
            vectors.append(TracingVector(
                file=file_name,
                sheet=sheet,
                cell_range=f"{_idx_to_col(start_c)}{row}:{_idx_to_col(end_c)}{row}",
                direction="row",
                length=len(vals),
                start_cell=f"{_idx_to_col(start_c)}{row}",
                end_cell=f"{_idx_to_col(end_c)}{row}",
                values=tuple(vals),
            ))

    return vectors


# ---------------------------------------------------------------------------
# Range helpers
# ---------------------------------------------------------------------------

def compute_sub_range(vec: TracingVector, offset: int, length: int) -> str:
    """Compute the cell range for a subsequence within a vector.

    *offset* is the 0-based index into the vector's values where the
    subsequence begins.
    """
    m = _CELL_RE.match(vec.start_cell)
    if not m:
        return vec.cell_range

    start_col_str = m.group(1)
    start_row = int(m.group(2))
    start_col = _col_to_idx(start_col_str)

    if vec.direction == "column":
        r1 = start_row + offset
        r2 = r1 + length - 1
        return f"{start_col_str}{r1}:{start_col_str}{r2}"
    else:
        c1 = start_col + offset
        c2 = c1 + length - 1
        return f"{_idx_to_col(c1)}{start_row}:{_idx_to_col(c2)}{start_row}"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def scan_model_sheet(
    path: Path,
    sheet_name: str,
    min_len: int = MIN_VECTOR_LEN,
) -> list[TracingVector]:
    """Scan one sheet of the model file for hardcoded numeric vectors.

    Returns TracingVectors with full values (not truncated to 5).
    """
    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            if sheet_name not in sheet_map:
                return []
            sheet_path = sheet_map[sheet_name]
            if sheet_path not in set(zf.namelist()):
                return []
            data = zf.read(sheet_path)
            cells = _stream_hardcoded_numerics(data)
            return _cells_to_vectors(cells, path.name, sheet_name, min_len)
    except Exception:
        return []


def scan_upstream_file(
    path: Path,
    min_len: int = MIN_VECTOR_LEN,
) -> list[TracingVector]:
    """Scan an upstream file for ALL numeric vectors (formula + hardcoded).

    Returns TracingVectors with full values for every sheet.
    Only supports XLSX/XLSM (ZIP-based).  Returns [] for XLS/XLSB.
    """
    all_vectors: list[TracingVector] = []
    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            names = set(zf.namelist())
            for sheet_name, sheet_path in sheet_map.items():
                if sheet_path not in names:
                    continue
                try:
                    data = zf.read(sheet_path)
                    cells = _stream_all_numerics(data)
                    vectors = _cells_to_vectors(cells, path.name, sheet_name, min_len)
                    all_vectors.extend(vectors)
                except Exception:
                    pass
    except zipfile.BadZipFile:
        pass  # XLS / XLSB — not supported
    except Exception:
        pass

    return all_vectors


def get_sheet_names(path: Path) -> list[str]:
    """Return the ordered list of sheet names from an Excel file."""
    try:
        with zipfile.ZipFile(path) as zf:
            return list(_get_sheet_map(zf).keys())
    except Exception:
        return []
