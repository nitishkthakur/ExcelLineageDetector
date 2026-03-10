"""Extractor for hardcoded (manually entered or copy-pasted) values in worksheets.

Finance analysts frequently bypass data connections by copying values from Bloomberg
terminals, Reuters, reports, or databases and pasting them directly into cells.
These break lineage — this extractor surfaces them as 'input' connections.

Detection logic:
1. Named ranges pointing to single value cells → named_input (highest confidence)
2. Cells in 'Input'/'Assumption' sheets with a row label → hardcoded_value
3. Labeled parameter rows (text label in col A + numeric value nearby) → hardcoded_value
4. Contiguous blocks of value cells with column headers → pasted_table
5. Text cells containing source attributions ("Source: Bloomberg") → source_note
"""

from __future__ import annotations
import re
import zipfile
from collections import defaultdict

from lxml import etree

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Sheet name signals that the whole sheet is an input/assumption sheet
INPUT_SHEET_RE = re.compile(
    r'\b(inputs?|assumptions?|params?|parameters?|drivers?|'
    r'hypothesis|hypotheses|key\s*inputs?|model\s*inputs?|data\s*sources?)\b',
    re.IGNORECASE,
)

# Source attribution patterns in cell text
SOURCE_ATTR_RE = re.compile(
    r'\b(source|per|from|data\s+from|as\s+of|per\s+bloomberg|'
    r'bloomberg|reuters|factset|capital\s*iq|morningstar|compustat|'
    r'refinitiv|wind\s+info|SNL|company\s+filings?|annual\s+report)\b',
    re.IGNORECASE,
)


# Column letters to 1-based index
def _col_to_idx(col: str) -> int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def _idx_to_col(idx: int) -> str:
    col = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        col = chr(rem + 65) + col
    return col


def _split_ref(ref: str):
    """Split 'AB12' -> ('AB', 12). Returns ('', 0) on failure."""
    m = re.match(r'^([A-Za-z]+)(\d+)$', ref.strip())
    if m:
        return m.group(1).upper(), int(m.group(2))
    return '', 0


class HardcodedValuesExtractor(BaseExtractor):
    """Surfaces hardcoded values (manually entered / copy-pasted) as 'input' connections."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            shared_strings = self._load_shared_strings(zip_file)
            sheet_map = self._get_sheet_map(zip_file)
            named_cells = self._load_named_value_cells(zip_file, shared_strings)

            for sheet_name, sheet_file in sheet_map.items():
                try:
                    found = self._extract_from_sheet(
                        zip_file, sheet_name, sheet_file,
                        shared_strings, named_cells,
                    )
                    connections.extend(found)
                except Exception as e:
                    self.log.warning(f"HardcodedValues: failed on {sheet_name}: {e}")
        except Exception as e:
            self.log.error(f"HardcodedValuesExtractor failed: {e}", exc_info=True)
        return connections

    # ------------------------------------------------------------------ sheet map

    def _get_sheet_map(self, zip_file: zipfile.ZipFile) -> dict[str, str]:
        sheet_map = {}
        try:
            wb_root = self._read_xml(zip_file, "xl/workbook.xml")
            if wb_root is None:
                return sheet_map
            REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            sheets = (wb_root.findall(f"{{{NS}}}sheets/{{{NS}}}sheet")
                      or wb_root.findall(".//sheets/sheet")
                      or wb_root.findall(".//{*}sheet"))
            rels = self._read_rels(zip_file, "xl/_rels/workbook.xml.rels")
            for sh in sheets:
                name = sh.get("name", "")
                rid = sh.get(f"{{{REL_NS}}}id") or sh.get("r:id") or sh.get("id", "")
                if rid in rels:
                    target = rels[rid]["target"].lstrip("/")
                    if not target.startswith("xl/"):
                        target = f"xl/{target}"
                    sheet_map[name] = target
        except Exception as e:
            self.log.warning(f"HardcodedValues: sheet map failed: {e}")
        if not sheet_map:
            for name in zip_file.namelist():
                if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                    idx = re.search(r"sheet(\d+)", name).group(1)
                    sheet_map[f"Sheet{idx}"] = name
        return sheet_map

    # ------------------------------------------------------------------ shared strings

    def _load_shared_strings(self, zip_file: zipfile.ZipFile) -> list[str]:
        strings: list[str] = []
        root = self._read_xml(zip_file, "xl/sharedStrings.xml")
        if root is None:
            return strings
        for si in (root.findall(f".//{{{NS}}}si")
                   or root.findall(".//si")
                   or root.findall(".//{*}si")):
            texts = []
            for t_el in (si.findall(f".//{{{NS}}}t")
                         or si.findall(".//t")
                         or si.findall(".//{*}t")):
                texts.append(t_el.text or "")
            strings.append("".join(texts))
        return strings

    # ------------------------------------------------------------------ named ranges

    def _load_named_value_cells(
        self, zip_file: zipfile.ZipFile, shared_strings: list[str]
    ) -> dict[str, str]:
        """Return map of cell_key -> named_range_name for named ranges that
        point to a single cell (not a multi-cell range or formula)."""
        named = {}
        root = self._read_xml(zip_file, "xl/workbook.xml")
        if root is None:
            return named
        for dn in (root.findall(f".//{{{NS}}}definedName")
                   or root.findall(".//definedName")
                   or root.findall(".//{*}definedName")):
            name = dn.get("name", "")
            formula = (dn.text or "").strip()
            # Single cell reference like Sheet1!$B$5 or 'Input Sheet'!$C$3
            m = re.match(r"'?([^'!]+)'?!\$?([A-Za-z]+)\$?(\d+)$", formula)
            if m and name:
                sheet = m.group(1).strip("'")
                col = m.group(2).upper()
                row = int(m.group(3))
                key = f"{sheet}!{col}{row}"
                named[key] = name
        return named

    # ------------------------------------------------------------------ sheet extraction

    def _extract_from_sheet(
        self,
        zip_file: zipfile.ZipFile,
        sheet_name: str,
        sheet_file: str,
        shared_strings: list[str],
        named_cells: dict[str, str],
    ) -> list[DataConnection]:
        results = []
        if sheet_file not in zip_file.namelist():
            return results

        try:
            data = zip_file.read(sheet_file)
            root = etree.fromstring(data)
        except Exception as e:
            self.log.warning(f"HardcodedValues: cannot parse {sheet_file}: {e}")
            return results

        is_input_sheet = bool(INPUT_SHEET_RE.search(sheet_name))

        # Parse all cells into a map: (row_idx, col_idx) -> cell_dict
        cells = self._parse_cells(root, shared_strings)

        if not cells:
            return results

        # Build row_labels: row_idx -> leftmost text in that row (up to col 4)
        row_labels: dict[int, str] = {}
        for (r, c), cell in cells.items():
            if (cell["type"] == "s"
                    and not cell["has_formula"]
                    and cell["value"]
                    and c <= 4):
                if r not in row_labels:
                    row_labels[r] = str(cell["value"]).strip()

        # col_headers: col_idx -> header text from row 1 or 2
        col_headers: dict[int, str] = {}
        for header_row in (1, 2):
            for (r, c), cell in cells.items():
                if r == header_row and cell["type"] == "s" and cell["value"] and c not in col_headers:
                    col_headers[c] = str(cell["value"]).strip()

        # Find source attribution text cells
        for (r, c), cell in cells.items():
            if cell["type"] == "s" and not cell["has_formula"] and cell["value"]:
                text = str(cell["value"])
                if SOURCE_ATTR_RE.search(text) and len(text) >= 8:
                    ref = cell["ref"]
                    loc = f"{sheet_name}!{ref}"
                    conn = DataConnection(
                        id=DataConnection.make_id("input", text[:80], loc),
                        category="input",
                        sub_type="source_note",
                        source=text[:100],
                        raw_connection=text,
                        location=loc,
                        metadata={
                            "sheet": sheet_name,
                            "cell": ref,
                            "label": "",
                            "value": text,
                            "value_type": "text",
                        },
                        confidence=0.7,
                    )
                    results.append(conn)

        # Find value (non-formula) cells
        value_cells = {
            (r, c): cell for (r, c), cell in cells.items()
            if not cell["has_formula"]
            and cell["value"] is not None
            and cell["type"] not in ("s", "b", "e")  # exclude strings, bools, errors
        }

        # Detect contiguous value regions (potential pasted tables)
        table_regions = self._detect_table_regions(value_cells, col_headers)
        cells_in_tables: set[tuple] = set()
        for region in table_regions:
            cells_in_tables.update(region["cells"])
            min_r = min(r for r, c in region["cells"])
            max_r = max(r for r, c in region["cells"])
            min_c = min(c for r, c in region["cells"])
            max_c = max(c for r, c in region["cells"])
            range_ref = (
                f"{_idx_to_col(min_c)}{min_r}:{_idx_to_col(max_c)}{max_r}"
            )
            headers_str = ", ".join(
                col_headers[c] for c in range(min_c, max_c + 1) if c in col_headers
            )
            loc = f"{sheet_name}!{range_ref}"
            conn = DataConnection(
                id=DataConnection.make_id("input", range_ref, loc),
                category="input",
                sub_type="pasted_table",
                source=headers_str[:100] if headers_str else range_ref,
                raw_connection=range_ref,
                location=loc,
                metadata={
                    "sheet": sheet_name,
                    "cell": range_ref,
                    "label": headers_str,
                    "value": range_ref,
                    "value_type": "table",
                    "rows": max_r - min_r + 1,
                    "cols": max_c - min_c + 1,
                    "col_headers": [col_headers.get(c, "") for c in range(min_c, max_c + 1)],
                },
                confidence=0.75,
            )
            results.append(conn)

        # Report individual value cells not in detected tables
        for (r, c), cell in value_cells.items():
            if (r, c) in cells_in_tables:
                continue
            ref = cell["ref"]
            cell_key = f"{sheet_name}!{ref}"

            # Check if named range
            named_range = named_cells.get(cell_key, "")

            row_label = row_labels.get(r, "")
            col_header = col_headers.get(c, "")
            value = cell["value"]
            value_str = str(value)

            # Determine confidence and whether to report
            confidence = 0.5
            if named_range:
                confidence = 0.95
            elif is_input_sheet and row_label:
                confidence = 0.9
            elif is_input_sheet:
                confidence = 0.7
            elif row_label and col_header:
                confidence = 0.85
            elif row_label:
                confidence = 0.75
            else:
                # No label context — skip to avoid noise
                continue

            label = named_range or row_label or col_header or ""
            sub_type = "named_input" if named_range else "hardcoded_value"

            loc = f"{sheet_name}!{ref}"
            conn = DataConnection(
                id=DataConnection.make_id("input", cell_key, loc),
                category="input",
                sub_type=sub_type,
                source=label[:100] if label else ref,
                raw_connection=value_str,
                location=loc,
                metadata={
                    "sheet": sheet_name,
                    "cell": ref,
                    "label": label,
                    "value": value_str,
                    "value_type": "number",
                    "named_range": named_range,
                    "col_header": col_header,
                },
                confidence=confidence,
            )
            results.append(conn)

        return results

    # ------------------------------------------------------------------ cell parsing

    def _parse_cells(self, root, shared_strings: list[str]) -> dict:
        """Parse all cells from a sheet XML root."""
        cells = {}
        row_els = (root.findall(f".//{{{NS}}}row")
                   or root.findall(".//row")
                   or root.findall(".//{*}row"))
        for row_el in row_els:
            row_idx_str = row_el.get("r", "0")
            try:
                row_idx = int(row_idx_str)
            except ValueError:
                continue

            c_els = (row_el.findall(f"{{{NS}}}c")
                     or row_el.findall("c")
                     or row_el.findall("{*}c"))
            for c_el in c_els:
                ref = c_el.get("r", "")
                if not ref:
                    continue
                col_str, _ = _split_ref(ref)
                if not col_str:
                    continue
                col_idx = _col_to_idx(col_str)

                cell_type = c_el.get("t", "n")
                f_el = c_el.find(f"{{{NS}}}f")
                if f_el is None:
                    f_el = c_el.find("f")
                if f_el is None:
                    f_el = c_el.find("{*}f")
                v_el = c_el.find(f"{{{NS}}}v")
                if v_el is None:
                    v_el = c_el.find("v")
                if v_el is None:
                    v_el = c_el.find("{*}v")
                is_el = c_el.find(f"{{{NS}}}is")
                if is_el is None:
                    is_el = c_el.find("is")
                if is_el is None:
                    is_el = c_el.find("{*}is")

                has_formula = f_el is not None

                value = None
                resolved_type = cell_type

                if v_el is not None and v_el.text is not None:
                    raw = v_el.text
                    if cell_type == "s":
                        try:
                            idx = int(raw)
                            value = shared_strings[idx] if 0 <= idx < len(shared_strings) else raw
                        except (ValueError, IndexError):
                            value = raw
                        resolved_type = "s"
                    elif cell_type == "b":
                        value = raw == "1"
                        resolved_type = "b"
                    elif cell_type in ("e", "str"):
                        value = raw
                        resolved_type = cell_type
                    else:
                        try:
                            fv = float(raw)
                            value = int(fv) if fv == int(fv) and abs(fv) < 1e15 else fv
                        except ValueError:
                            value = raw
                        resolved_type = "n"
                elif is_el is not None:
                    t = is_el.find(f"{{{NS}}}t")
                    if t is None:
                        t = is_el.find("t")
                    if t is None:
                        t = is_el.find("{*}t")
                    if t is not None:
                        value = t.text or ""
                    resolved_type = "s"

                cells[(row_idx, col_idx)] = {
                    "ref": ref,
                    "row": row_idx,
                    "col": col_idx,
                    "value": value,
                    "has_formula": has_formula,
                    "type": resolved_type,
                }
        return cells

    # ------------------------------------------------------------------ table detection

    def _detect_table_regions(
        self,
        value_cells: dict,
        col_headers: dict,
    ) -> list[dict]:
        """Detect contiguous rectangular blocks of value cells (pasted tables).

        A table region must be:
        - At least 2 rows x 2 columns of value cells
        - Have at least 2 column headers (row 1/2) above it
        """
        if not value_cells:
            return []

        visited = set()
        regions = []

        for (r, c) in sorted(value_cells.keys()):
            if (r, c) in visited:
                continue

            # BFS to find contiguous block
            block = set()
            queue = [(r, c)]
            while queue:
                cr, cc = queue.pop()
                if (cr, cc) in block or (cr, cc) not in value_cells:
                    continue
                block.add((cr, cc))
                for dr, dc in ((0, 1), (0, -1), (1, 0), (-1, 0)):
                    nb = (cr + dr, cc + dc)
                    if nb in value_cells and nb not in block:
                        queue.append(nb)

            visited.update(block)

            if len(block) < 4:  # need at least 2x2
                continue

            rows_in_block = sorted(set(r for r, c in block))
            cols_in_block = sorted(set(c for r, c in block))

            if len(rows_in_block) < 2 or len(cols_in_block) < 2:
                continue

            # Check that at least 2 cols have headers
            headers_present = sum(1 for c in cols_in_block if c in col_headers)
            if headers_present < 2:
                continue

            regions.append({"cells": block})

        return regions
