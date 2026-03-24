"""Excel report writer for RawSourcesDetection results.

Produces a 6-sheet workbook:
  1. Summary              — run settings and stats
  2. Required Inputs      — full dump of every upstream source that must be supplied
  3. Formula Dependency   — all external formula references by level
  4. Missing Files        — referenced files not found on disk
  5. Matched Vectors      — hardcoded vectors with confirmed source
  6. Unmatched Vectors    — hardcoded vectors with no source found
"""
from __future__ import annotations

from pathlib import Path

from .config import RSDConfig
from .models import DetectionResult


# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
_HDR_FILL    = "1565C0"   # dark blue — header background
_HDR_TEXT    = "FFFFFF"   # white — header text
_FOUND_FILL  = "C8E6C9"   # green — file found on disk
_MISSING_FILL = "FFCDD2"  # red/salmon — file not found
_MISS_ALT    = "FFEBEE"   # lighter red — alternating missing row
_EXACT_FILL  = "C8E6C9"   # green — exact match
_SUBSEQ_FILL = "DCEDC8"   # light green — exact subsequence
_APPROX_FILL = "BBDEFB"   # blue — approximate match
_UNMATCH_FILL = "FFF9C4"  # yellow — unmatched vector
_UNMATCH_ALT  = "FFFDE7"  # lighter yellow — alternating
_DB_FILL     = "E3F2FD"   # blue — database / ODBC / OLE DB
_FILE_FILL   = "E8F5E9"   # green — file reference
_PQ_FILL     = "FFF3E0"   # orange — Power Query
_OTHER_FILL  = "F3E5F5"   # purple — other categories


# Map category → background colour for Required Inputs sheet
_CAT_FILL = {
    "database":   _DB_FILL,
    "powerquery": _PQ_FILL,
    "file":       _FILE_FILL,
    "formula":    _FILE_FILL,
}

# Map match_type → background colour for Matched Vectors sheet
_MATCH_FILL = {
    "exact":             _EXACT_FILL,
    "exact_subsequence": _SUBSEQ_FILL,
    "approximate":       _APPROX_FILL,
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fmt_sample(values: list[float]) -> str:
    """Format first 5 sample values compactly."""
    parts = []
    for v in values[:5]:
        if isinstance(v, float) and v == int(v) and abs(v) < 1e12:
            parts.append(str(int(v)))
        else:
            parts.append(f"{v:.4g}")
    return ", ".join(parts)


# ---------------------------------------------------------------------------
# Main writer
# ---------------------------------------------------------------------------

def write_report(
    result: DetectionResult,
    config: RSDConfig,
    model_path: Path,
    out_path: Path,
) -> None:
    """Write the full DetectionResult to an Excel workbook at *out_path*."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError as e:
        raise ImportError(f"openpyxl required: {e}")

    wb = Workbook()
    wb.remove(wb.active)

    # ── Shared style factories ──────────────────────────────────────────────

    hdr_fill  = PatternFill("solid", fgColor=_HDR_FILL)
    hdr_font  = Font(bold=True, color=_HDR_TEXT, size=11)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="CCCCCC")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    def pfill(hex_color: str) -> PatternFill:
        return PatternFill("solid", fgColor=hex_color)

    def hdr(ws, headers: list[str], widths: list[int] | None = None) -> None:
        """Write header row with blue fill."""
        ws.row_dimensions[1].height = 30
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(1, ci, h)
            cell.fill  = hdr_fill
            cell.font  = hdr_font
            cell.alignment = hdr_align
            cell.border = border
        if widths:
            for ci, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(ci)].width = w

    def cell(ws, row: int, col: int, value, fill=None,
             bold: bool = False, wrap: bool = False):
        """Write a data cell with consistent styling."""
        c = ws.cell(row, col, value)
        c.border = border
        c.font   = Font(bold=bold, size=10)
        c.alignment = Alignment(wrap_text=wrap, vertical="top")
        if fill:
            c.fill = fill
        return c

    def autowidth(ws, min_w: int = 10, max_w: int = 50) -> None:
        """Auto-size column widths, clamped to [min_w, max_w]."""
        for col_cells in ws.columns:
            best = max(
                (len(str(c.value)) for c in col_cells if c.value is not None),
                default=0,
            )
            ci = col_cells[0].column
            ws.column_dimensions[get_column_letter(ci)].width = max(min_w, min(max_w, best + 2))

    # ========================================================================
    # Sheet 1: Summary
    # ========================================================================
    ws = wb.create_sheet("Summary")
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 50

    ws["A1"] = "RawSourcesDetection Report"
    ws["A1"].font = Font(bold=True, size=14, color="1565C0")
    ws.merge_cells("A1:B1")

    n_found = sum(1 for n in result.source_nodes if n.found_on_disk and n.level > 0)

    rows: list[tuple[str, object]] = [
        ("", ""),
        ("Model File", result.model_file),
        ("Sheets Traced", ", ".join(config.model_sheets) if config.model_sheets else "(all)"),
        ("Max Formula Levels", config.max_formula_levels),
        ("Match Mode", "Exact + Approximate" if config.approximate else "Exact only"),
        ("Decimal Places (Exact)", config.exact_decimal_places),
        ("Subsequence Matching", "Yes" if config.subsequence_matching else "No"),
        ("Min Vector Length", config.min_vector_length),
        ("", ""),
        ("── Formula Tracing ──", ""),
        ("Files Found on Disk", n_found),
        ("Files Missing (not on disk)", len(result.missing_files)),
        ("Formula Refs Found", len(result.formula_refs)),
        ("Formula Levels Reached", max((r.level for r in result.formula_refs), default=0)),
        ("", ""),
        ("── Vector Matching ──", ""),
        ("Vectors Matched", len(result.matched_vectors)),
        ("Vectors Unmatched", len(result.unmatched_vectors)),
        ("", ""),
        ("── Connection Harvesting ──", ""),
        ("Raw Sources Found", len(result.raw_sources)),
        ("  — Database / ODBC / OLE DB",
         sum(1 for s in result.raw_sources if s.category == "database")),
        ("  — Power Query",
         sum(1 for s in result.raw_sources if s.category == "powerquery")),
        ("  — File References",
         sum(1 for s in result.raw_sources if s.category == "file")),
        ("  — Other",
         sum(1 for s in result.raw_sources
             if s.category not in ("database", "powerquery", "file"))),
    ]

    lbl_font = Font(bold=True, size=10)
    val_font = Font(size=10)
    for ri, (label, value) in enumerate(rows, 2):
        ws.cell(ri, 1, label).font = lbl_font
        ws.cell(ri, 2, str(value) if value != "" else "").font = val_font

    # ========================================================================
    # Sheet 2: Required Inputs (pure dump — "shopping list")
    # ========================================================================
    ws = wb.create_sheet("Required Inputs")
    hdr(ws,
        ["Type", "Sub-Type", "Connection / Path / File",
         "Found On Disk", "Discovered In", "Location / Cell"],
        [14, 16, 52, 15, 26, 32])

    ri = 2
    seen_req: set[str] = set()

    # Part A: Excel files from formula tracing (deduplicated by filename)
    for ref in result.formula_refs:
        key = ref.target_file.lower()
        if key in seen_req:
            continue
        seen_req.add(key)
        f = pfill(_FOUND_FILL if ref.file_found else _MISSING_FILL)
        cell(ws, ri, 1, "Excel File", f)
        cell(ws, ri, 2, "formula_ref", f)
        cell(ws, ri, 3, ref.target_file, f)
        cell(ws, ri, 4, "Yes" if ref.file_found else "NO — MISSING", f, bold=not ref.file_found)
        cell(ws, ri, 5, ref.source_file, f)
        cell(ws, ri, 6, f"{ref.source_sheet}!{ref.source_cell}" if ref.source_cell else ref.source_sheet, f)
        ri += 1

    # Part B: ODBC / OLE DB / Power Query / other connections from all files
    for src in result.raw_sources:
        f = pfill(_CAT_FILL.get(src.category, _OTHER_FILL))
        cell(ws, ri, 1, src.category, f)
        cell(ws, ri, 2, src.sub_type, f)
        cell(ws, ri, 3, src.connection, f, wrap=True)
        cell(ws, ri, 4, "N/A", f)
        cell(ws, ri, 5, src.source_file, f)
        cell(ws, ri, 6, src.location, f)
        ri += 1

    autowidth(ws)

    # ========================================================================
    # Sheet 3: Formula Dependency Tree
    # ========================================================================
    ws = wb.create_sheet("Formula Dependency Tree")
    hdr(ws,
        ["Level", "Source File", "Source Sheet", "Source Cell",
         "Target File", "Target Sheet", "Target Range",
         "Found on Disk", "Resolved Path"],
        [8, 24, 20, 12, 24, 20, 18, 14, 42])

    for ri, ref in enumerate(result.formula_refs, 2):
        f = pfill(_FOUND_FILL if ref.file_found else _MISSING_FILL)
        cell(ws, ri, 1, ref.level, f)
        cell(ws, ri, 2, ref.source_file, f)
        cell(ws, ri, 3, ref.source_sheet, f)
        cell(ws, ri, 4, ref.source_cell, f)
        cell(ws, ri, 5, ref.target_file, f, bold=not ref.file_found)
        cell(ws, ri, 6, ref.target_sheet, f)
        cell(ws, ri, 7, ref.target_range, f)
        cell(ws, ri, 8, "Yes" if ref.file_found else "NO", f)
        cell(ws, ri, 9, ref.resolved_path, f)

    autowidth(ws)

    # ========================================================================
    # Sheet 4: Missing Files
    # ========================================================================
    ws = wb.create_sheet("Missing Files")
    hdr(ws,
        ["Missing File", "Level", "Referenced By",
         "Sheets Needed", "Cells Referencing (first 10)"],
        [32, 8, 28, 32, 44])

    for ri, mf in enumerate(result.missing_files, 2):
        f = pfill(_MISSING_FILL if ri % 2 == 0 else _MISS_ALT)
        cell(ws, ri, 1, mf.filename, f, bold=True)
        cell(ws, ri, 2, mf.level, f)
        cell(ws, ri, 3, mf.referenced_by, f)
        cell(ws, ri, 4, ", ".join(mf.sheets_needed), f, wrap=True)
        refs_text = "\n".join(mf.cells_referencing[:10])
        cell(ws, ri, 5, refs_text, f, wrap=True)
        if mf.cells_referencing:
            ws.row_dimensions[ri].height = min(15 * min(len(mf.cells_referencing), 10), 80)

    autowidth(ws)

    # ========================================================================
    # Sheet 5: Matched Vectors
    # ========================================================================
    ws = wb.create_sheet("Matched Vectors")
    hdr(ws,
        ["Model Sheet", "Model Range", "Length", "Model Sample (first 5)",
         "Match Type", "Similarity",
         "Upstream File", "Upstream Sheet", "Upstream Range", "Upstream Sample"],
        [18, 14, 8, 30, 18, 10, 28, 18, 16, 30])

    for ri, mv in enumerate(result.matched_vectors, 2):
        f = pfill(_MATCH_FILL.get(mv.match_type, _EXACT_FILL))
        cell(ws, ri, 1, mv.model_sheet, f)
        cell(ws, ri, 2, mv.model_range, f)
        cell(ws, ri, 3, mv.model_length, f)
        cell(ws, ri, 4, _fmt_sample(mv.model_sample), f)
        cell(ws, ri, 5, mv.match_type, f)
        cell(ws, ri, 6, round(mv.similarity, 6), f)
        cell(ws, ri, 7, mv.upstream_file, f)
        cell(ws, ri, 8, mv.upstream_sheet, f)
        cell(ws, ri, 9, mv.upstream_range, f)
        cell(ws, ri, 10, _fmt_sample(mv.upstream_sample), f)

    autowidth(ws)

    # ========================================================================
    # Sheet 6: Unmatched Vectors
    # ========================================================================
    ws = wb.create_sheet("Unmatched Vectors")
    hdr(ws,
        ["Model Sheet", "Model Range", "Length",
         "Sample Values (first 5)", "Action Required"],
        [18, 14, 8, 36, 48])

    for ri, uv in enumerate(result.unmatched_vectors, 2):
        f = pfill(_UNMATCH_FILL if ri % 2 == 0 else _UNMATCH_ALT)
        cell(ws, ri, 1, uv.model_sheet, f)
        cell(ws, ri, 2, uv.model_range, f)
        cell(ws, ri, 3, uv.model_length, f)
        cell(ws, ri, 4, _fmt_sample(uv.model_sample), f)
        cell(ws, ri, 5, "No source found — verify manually or supply upstream file", f)

    autowidth(ws)

    # ── Save ────────────────────────────────────────────────────────────────
    wb.save(str(out_path))
