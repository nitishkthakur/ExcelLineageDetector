"""MCP server exposing Excel reading tools for LLM business-name inference."""
from __future__ import annotations

import threading
import zipfile
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from .streaming import (
    CellInfo,
    _load_shared_strings,
    get_named_ranges,
    get_sheet_summary,
    index_to_col_letter,
    list_sheets,
    read_cell_neighborhood,
    stream_sheet_cells,
)

# Thread-safe state with a lock
_lock = threading.Lock()
_EXCEL_PATH: Path | None = None
_ZF: zipfile.ZipFile | None = None
_SHARED_STRINGS: list[str] | None = None

mcp = FastMCP("excel-reader")


def configure(excel_path: Path) -> None:
    """Set the Excel file the server will expose. Thread-safe."""
    global _EXCEL_PATH, _ZF, _SHARED_STRINGS
    with _lock:
        if _ZF is not None:
            _ZF.close()
        _EXCEL_PATH = excel_path
        _ZF = zipfile.ZipFile(str(excel_path), "r")
        _SHARED_STRINGS = _load_shared_strings(_ZF)


def cleanup() -> None:
    """Close the zip file. Thread-safe."""
    global _ZF, _SHARED_STRINGS
    with _lock:
        if _ZF:
            _ZF.close()
            _ZF = None
        _SHARED_STRINGS = None


def _ensure_open() -> tuple[zipfile.ZipFile, list[str]]:
    """Return (zf, shared_strings), raising if not configured."""
    with _lock:
        if _ZF is None:
            raise RuntimeError("Server not configured — call configure() first")
        return _ZF, _SHARED_STRINGS or []


@mcp.tool()
def read_cell_neighborhood_tool(
    sheet_name: str,
    center_ref: str,
    radius: int = 3,
) -> dict:
    """Read cells within +/-radius rows/cols of a center cell.

    Args:
        sheet_name: Name of the sheet (e.g. "Inputs")
        center_ref: Cell reference (e.g. "B3")
        radius: Number of rows/cols around center to include (default 3)

    Returns a dict mapping cell refs to their values/formulas.
    """
    zf, ss = _ensure_open()
    cells = read_cell_neighborhood(zf, sheet_name, center_ref, radius, ss)
    return {
        ref: {
            "value": c.value,
            "formula": c.formula,
            "type": c.cell_type,
        }
        for ref, c in cells.items()
    }


@mcp.tool()
def get_sheet_summary_tool(
    sheet_name: str,
    max_rows: int = 5,
) -> dict:
    """Get a summary of a sheet: dimensions, headers, and sample rows.

    Args:
        sheet_name: Name of the sheet
        max_rows: Number of sample data rows to return (default 5)
    """
    zf, ss = _ensure_open()
    return get_sheet_summary(zf, sheet_name, ss, max_rows)


@mcp.tool()
def get_headers_for_range(
    sheet_name: str,
    start_col: str,
    end_col: str,
) -> list[str | None]:
    """Get the header row (row 1) values for a column range.

    Args:
        sheet_name: Name of the sheet
        start_col: Start column letter (e.g. "B")
        end_col: End column letter (e.g. "M")

    Returns list of header values, one per column.
    """
    zf, ss = _ensure_open()
    from .streaming import col_letter_to_index

    start_idx = col_letter_to_index(start_col)
    end_idx = col_letter_to_index(end_col)

    headers: dict[int, str | None] = {}
    for cell in stream_sheet_cells(zf, sheet_name, ss):
        if cell.row == 1 and start_idx <= cell.col <= end_idx:
            headers[cell.col] = cell.value
        if cell.row > 1:
            break

    return [headers.get(c) for c in range(start_idx, end_idx + 1)]


@mcp.tool()
def list_sheets_tool() -> list[str]:
    """List all sheet names in the Excel file."""
    zf, _ = _ensure_open()
    return list_sheets(zf)


@mcp.tool()
def get_named_ranges_tool() -> list[dict]:
    """Get all defined names / named ranges in the workbook.

    Returns list of dicts with name, value, scope_sheet_id, hidden.
    """
    zf, _ = _ensure_open()
    return get_named_ranges(zf)
