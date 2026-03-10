"""Excel reporter for Excel Lineage Detector."""

from __future__ import annotations
from collections import Counter
from pathlib import Path
from typing import Optional

from lineage.models import DataConnection


# Category color coding (background fill colors).
# Add an entry here whenever a new category is introduced.
CATEGORY_COLORS = {
    "database": "E3F2FD",
    "file":     "E8F5E9",
    "web":      "FFF3E0",
    "powerquery": "F3E5F5",
    "vba":      "FFEBEE",
    "pivot":    "E0F7FA",
    "formula":  "E8EAF6",
    "hyperlink": "FFF8E1",
    "ole":      "F9FBE7",
    "metadata": "FAFAFA",
    "input":    "FFF9C4",  # light yellow - signals manual entry
}

# Categories that have their own dedicated sheet in the report.
# All other categories land in the "Other Sources" catch-all sheet.
DEDICATED_SHEET_CATEGORIES = frozenset(
    ["database", "powerquery", "file", "web", "hyperlink", "vba", "input"]
)

HEADER_FILL_COLOR = "1565C0"
HEADER_TEXT_COLOR = "FFFFFF"


class ExcelReporter:
    """Creates a formatted Excel report of lineage findings."""

    def write(
        self,
        connections: list[DataConnection],
        input_path: Path,
        out_dir: Path,
    ) -> Path:
        """Write connections to a formatted Excel report.

        Args:
            connections: List of detected DataConnection objects.
            input_path: Path to the analyzed Excel file.
            out_dir: Directory to write the output file.

        Returns:
            Path to the written Excel file.
        """
        try:
            from openpyxl import Workbook
            from openpyxl.styles import (
                PatternFill, Font, Alignment, Border, Side, numbers
            )
            from openpyxl.chart import BarChart, Reference
            from openpyxl.utils import get_column_letter
        except ImportError as e:
            raise ImportError(f"openpyxl is required for Excel reporting: {e}")

        stem = input_path.stem
        out = out_dir / f"{stem}_lineage_report.xlsx"

        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet

        # Create header style helpers
        def make_header_fill():
            return PatternFill("solid", fgColor=HEADER_FILL_COLOR)

        def make_header_font():
            return Font(bold=True, color=HEADER_TEXT_COLOR, size=11)

        def make_category_fill(category: str):
            color = CATEGORY_COLORS.get(category, "FFFFFF")
            return PatternFill("solid", fgColor=color)

        def make_border():
            side = Side(style="thin", color="CCCCCC")
            return Border(left=side, right=side, top=side, bottom=side)

        def make_header_border():
            side = Side(style="medium", color="FFFFFF")
            return Border(left=side, right=side, top=side, bottom=side)

        def style_header_row(ws, headers: list[str]):
            """Apply header styles to first row."""
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.fill = make_header_fill()
                cell.font = make_header_font()
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = make_header_border()
            ws.row_dimensions[1].height = 22

        def auto_width(ws, min_width=10, max_width=60):
            """Auto-size column widths based on content."""
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    except Exception:
                        pass
                adjusted_width = min(max(max_len + 2, min_width), max_width)
                ws.column_dimensions[col_letter].width = adjusted_width

        # --- Sheet 1: Summary ---
        ws_summary = wb.create_sheet("Summary")
        self._write_summary_sheet(ws_summary, connections, input_path,
                                  style_header_row, make_header_fill,
                                  make_header_font, auto_width)

        # --- Sheet 2: All Connections ---
        ws_all = wb.create_sheet("All Connections")
        self._write_all_connections_sheet(ws_all, connections,
                                          style_header_row, make_category_fill,
                                          make_border, auto_width)

        # --- Sheet 3: Databases ---
        ws_db = wb.create_sheet("Databases")
        db_conns = [c for c in connections if c.category == "database"]
        self._write_database_sheet(ws_db, db_conns, style_header_row,
                                   make_category_fill, make_border, auto_width)

        # --- Sheet 4: Power Query ---
        ws_pq = wb.create_sheet("Power Query")
        pq_conns = [c for c in connections if c.category == "powerquery"]
        self._write_powerquery_sheet(ws_pq, pq_conns, style_header_row,
                                     make_category_fill, make_border, auto_width)

        # --- Sheet 5: Files ---
        ws_files = wb.create_sheet("Files")
        file_conns = [c for c in connections if c.category == "file"]
        self._write_files_sheet(ws_files, file_conns, style_header_row,
                                make_category_fill, make_border, auto_width)

        # --- Sheet 6: Web & API ---
        ws_web = wb.create_sheet("Web & API")
        web_conns = [c for c in connections if c.category in ("web", "hyperlink")]
        self._write_web_sheet(ws_web, web_conns, style_header_row,
                              make_category_fill, make_border, auto_width)

        # --- Sheet 7: VBA ---
        ws_vba = wb.create_sheet("VBA")
        vba_conns = [c for c in connections if c.category == "vba"]
        self._write_vba_sheet(ws_vba, vba_conns, style_header_row,
                              make_category_fill, make_border, auto_width)

        # --- Sheet 8: Hardcoded Inputs ---
        input_conns = [c for c in connections if c.category == "input"]
        if input_conns:
            ws_inputs = wb.create_sheet("Hardcoded Inputs")
            self._write_inputs_sheet(ws_inputs, input_conns, style_header_row,
                                     make_category_fill, make_border, auto_width)

        # --- Sheet 9: Other Sources ---
        # Catches pivot, formula, ole, metadata, and any future new categories
        # so they are never silently missing from the report.
        other_conns = [c for c in connections
                       if c.category not in DEDICATED_SHEET_CATEGORIES]
        if other_conns:
            ws_other = wb.create_sheet("Other Sources")
            self._write_other_sheet(ws_other, other_conns, style_header_row,
                                    make_category_fill, make_border, auto_width)

        wb.save(str(out))
        return out

    def _write_summary_sheet(self, ws, connections, input_path,
                              style_header_row, make_header_fill,
                              make_header_font, auto_width):
        """Write summary statistics sheet."""
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.chart import BarChart, Reference

        ws.title = "Summary"
        row = 1

        # Title
        ws.cell(row=row, column=1, value="Excel Lineage Detector Report")
        ws.cell(row=row, column=1).font = Font(bold=True, size=16, color="1565C0")
        row += 1

        ws.cell(row=row, column=1, value=f"File: {input_path}")
        ws.cell(row=row, column=1).font = Font(size=10)
        row += 1

        from datetime import datetime
        ws.cell(row=row, column=1, value=f"Scanned: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
        row += 1
        ws.cell(row=row, column=1, value=f"Total Connections: {len(connections)}")
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        row += 2

        # Category table
        headers = ["Category", "Count", "% of Total"]
        style_header_row(ws, headers)
        ws.append([])  # shift down by re-doing rows
        # Actually write manually at current row
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col_idx, value=h)
            cell.fill = make_header_fill()
            cell.font = make_header_font()
            cell.alignment = Alignment(horizontal="center")
        row += 1

        by_cat = Counter(c.category for c in connections)
        total = len(connections)
        chart_start_row = row

        for cat, count in sorted(by_cat.items(), key=lambda x: -x[1]):
            pct = f"{count/total*100:.1f}%" if total > 0 else "0%"
            ws.cell(row=row, column=1, value=cat)
            ws.cell(row=row, column=2, value=count)
            ws.cell(row=row, column=3, value=pct)
            row += 1

        # Add bar chart
        if by_cat:
            chart = BarChart()
            chart.type = "col"
            chart.title = "Connections by Category"
            chart.y_axis.title = "Count"
            chart.x_axis.title = "Category"
            chart.style = 10
            chart.width = 20
            chart.height = 12

            data_ref = Reference(ws, min_col=2, min_row=chart_start_row - 1,
                                 max_row=chart_start_row + len(by_cat) - 1)
            cats_ref = Reference(ws, min_col=1, min_row=chart_start_row,
                                 max_row=chart_start_row + len(by_cat) - 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            ws.add_chart(chart, f"E6")

        auto_width(ws)

    def _write_all_connections_sheet(self, ws, connections,
                                     style_header_row, make_category_fill,
                                     make_border, auto_width):
        """Write all connections sheet with filtering."""
        from openpyxl.styles import Alignment

        headers = ["ID", "Category", "Sub-type", "Source", "Location",
                   "Query/Formula", "Confidence", "Raw Connection"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        for i, conn in enumerate(connections, 2):
            fill = make_category_fill(conn.category)
            row_data = [
                conn.id,
                conn.category,
                conn.sub_type,
                conn.source,
                conn.location,
                (conn.query_text or "")[:200] if conn.query_text else "",
                round(conn.confidence, 2),
                conn.raw_connection[:200],
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()
                if col_idx == 6:  # Query column - wrap text
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

        # Add auto-filter
        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        auto_width(ws)

        # Set specific column widths
        ws.column_dimensions["A"].width = 18  # ID
        ws.column_dimensions["D"].width = 35  # Source
        ws.column_dimensions["E"].width = 30  # Location
        ws.column_dimensions["F"].width = 45  # Query
        ws.column_dimensions["H"].width = 40  # Raw

    def _write_database_sheet(self, ws, connections,
                               style_header_row, make_category_fill,
                               make_border, auto_width):
        """Write databases sheet with SQL details."""
        from openpyxl.styles import Alignment

        headers = ["ID", "Sub-type", "Source", "Location",
                   "Tables", "Columns", "Query Text"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        fill = make_category_fill("database")
        for i, conn in enumerate(connections, 2):
            tables = ""
            columns = ""
            if conn.parsed_query:
                tables = ", ".join(conn.parsed_query.tables)
                columns = ", ".join(conn.parsed_query.columns)

            row_data = [
                conn.id, conn.sub_type, conn.source, conn.location,
                tables, columns,
                (conn.query_text or "")[:300],
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()
                if col_idx == 7:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        auto_width(ws)

    def _write_powerquery_sheet(self, ws, connections,
                                 style_header_row, make_category_fill,
                                 make_border, auto_width):
        """Write Power Query sheet with M code."""
        from openpyxl.styles import Alignment, Font

        headers = ["ID", "Query Name", "Sub-type", "Source", "Location", "M Formula"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        fill = make_category_fill("powerquery")
        mono_font = Font(name="Courier New", size=9)

        for i, conn in enumerate(connections, 2):
            query_name = conn.metadata.get("query_name", conn.source) if conn.metadata else conn.source
            row_data = [
                conn.id, query_name, conn.sub_type,
                conn.source, conn.location,
                (conn.query_text or "")[:500],
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()
                if col_idx == 6:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    cell.font = mono_font

        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        ws.column_dimensions["F"].width = 60
        auto_width(ws)
        ws.row_dimensions[1].height = 22

    def _write_files_sheet(self, ws, connections,
                            style_header_row, make_category_fill,
                            make_border, auto_width):
        """Write files sheet."""
        headers = ["ID", "Sub-type", "File Path", "Location", "Confidence"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        fill = make_category_fill("file")
        for i, conn in enumerate(connections, 2):
            row_data = [
                conn.id, conn.sub_type, conn.raw_connection,
                conn.location, round(conn.confidence, 2),
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()

        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        auto_width(ws)

    def _write_web_sheet(self, ws, connections,
                          style_header_row, make_category_fill,
                          make_border, auto_width):
        """Write web & API sheet."""
        headers = ["ID", "Category", "Sub-type", "URL", "Location", "Confidence"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        for i, conn in enumerate(connections, 2):
            fill = make_category_fill(conn.category)
            row_data = [
                conn.id, conn.category, conn.sub_type,
                conn.raw_connection, conn.location, round(conn.confidence, 2),
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()

        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        auto_width(ws)

    def _write_vba_sheet(self, ws, connections,
                          style_header_row, make_category_fill,
                          make_border, auto_width):
        """Write VBA sheet with code snippets."""
        from openpyxl.styles import Alignment, Font

        headers = ["ID", "Sub-type", "Source", "Location", "Code Snippet"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        fill = make_category_fill("vba")
        mono_font = Font(name="Courier New", size=9)

        for i, conn in enumerate(connections, 2):
            snippet = (conn.query_text or conn.raw_connection or "")[:300]
            row_data = [
                conn.id, conn.sub_type, conn.source,
                conn.location, snippet,
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()
                if col_idx == 5:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    cell.font = mono_font

        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        ws.column_dimensions["E"].width = 60
        auto_width(ws)

    def _write_inputs_sheet(self, ws, connections,
                             style_header_row, make_category_fill,
                             make_border, auto_width):
        """Write hardcoded inputs / manual values sheet."""
        from openpyxl.styles import Alignment

        headers = ["ID", "Sub-type", "Sheet", "Cell", "Label / Name",
                   "Value", "Value Type", "Confidence"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        fill = make_category_fill("input")
        for i, conn in enumerate(connections, 2):
            meta = conn.metadata or {}
            row_data = [
                conn.id,
                conn.sub_type,
                meta.get("sheet", conn.location.split("!")[0] if "!" in conn.location else ""),
                meta.get("cell", ""),
                meta.get("label", conn.source),
                meta.get("value", conn.raw_connection)[:200],
                meta.get("value_type", ""),
                round(conn.confidence, 2),
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()

        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        ws.column_dimensions["E"].width = 35  # Label
        ws.column_dimensions["F"].width = 35  # Value
        auto_width(ws)

    def _write_other_sheet(self, ws, connections,
                            style_header_row, make_category_fill,
                            make_border, auto_width):
        """Catch-all sheet for pivot, ole, formula, metadata, and any future categories.

        This sheet is intentionally generic so it works for every category without
        requiring code changes when new extractors are added.
        """
        from openpyxl.styles import Alignment

        headers = ["ID", "Category", "Sub-type", "Source", "Location",
                   "Raw Connection", "Query / Detail", "Confidence"]
        style_header_row(ws, headers)
        ws.freeze_panes = "A2"

        for i, conn in enumerate(connections, 2):
            fill = make_category_fill(conn.category)
            detail = (conn.query_text or "")[:300] if conn.query_text else ""
            row_data = [
                conn.id,
                conn.category,
                conn.sub_type,
                conn.source,
                conn.location,
                conn.raw_connection[:200],
                detail,
                round(conn.confidence, 2),
            ]
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=i, column=col_idx, value=value)
                cell.fill = fill
                cell.border = make_border()
                if col_idx == 7:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

        ws.auto_filter.ref = f"A1:{chr(ord('A') + len(headers) - 1)}1"
        ws.column_dimensions["F"].width = 40  # Raw Connection
        ws.column_dimensions["G"].width = 50  # Query / Detail
        auto_width(ws)
