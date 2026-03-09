"""Extractor for external references in cell formulas."""

from __future__ import annotations
import re
import zipfile

from lxml import etree

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection
from lineage.parsers.formula_parser import parse as parse_formula


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Compile regex patterns once
# External workbook reference: [workbook.xlsx] or path\[workbook.xlsx]
EXTERNAL_WB_PATTERN = re.compile(
    r"'?(?:[A-Za-z]:\\[^'\[]*|\\\\[^'\[]*)?(?:\[([^\]]+\.(?:xlsx?|xlsm|xlsb|csv))\])",
    re.IGNORECASE,
)

# UNC path in formula
UNC_PATTERN = re.compile(r"'(\\\\[^'\[]+)\[", re.IGNORECASE)

# Local drive path in formula
LOCAL_PATH_PATTERN = re.compile(r"'([A-Za-z]:\\[^'\[]+)\[", re.IGNORECASE)

# WEBSERVICE function
WEBSERVICE_PATTERN = re.compile(
    r'(?i)WEBSERVICE\s*\(\s*["\']?([^"\')\s,]+)',
    re.IGNORECASE,
)

# RTD function
RTD_PATTERN = re.compile(r'(?i)\bRTD\s*\(', re.IGNORECASE)

# FILTERXML with WEBSERVICE
FILTERXML_WS_PATTERN = re.compile(
    r'(?i)FILTERXML\s*\(\s*WEBSERVICE\s*\(\s*["\']?([^"\')\s,]+)',
    re.IGNORECASE,
)


class FormulasExtractor(BaseExtractor):
    """Extracts external references from cell formulas."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            sheet_map = self._get_sheet_map(zip_file)
            for sheet_name, sheet_file in sheet_map.items():
                try:
                    found = self._extract_from_sheet(zip_file, sheet_name, sheet_file)
                    connections.extend(found)
                except Exception as e:
                    self.log.warning(f"Failed to extract formulas from {sheet_name}: {e}")
        except Exception as e:
            self.log.error(f"FormulasExtractor failed: {e}", exc_info=True)
        return connections

    def _get_sheet_map(self, zip_file: zipfile.ZipFile) -> dict[str, str]:
        """Get mapping of sheet name -> sheet file path."""
        sheet_map = {}
        try:
            wb_root = self._read_xml(zip_file, "xl/workbook.xml")
            if wb_root is None:
                return sheet_map

            # Try both with and without namespace
            sheets = wb_root.findall(f"{{{NS}}}sheets/{{{NS}}}sheet")
            if not sheets:
                sheets = wb_root.findall(".//sheets/sheet")
            if not sheets:
                sheets = wb_root.findall(".//{*}sheet")

            # Get relationship mapping
            rels = self._read_rels(zip_file, "xl/_rels/workbook.xml.rels")

            for sheet_el in sheets:
                name = sheet_el.get("name", "")
                rid = sheet_el.get(f"{{{REL_NS}}}id") or sheet_el.get("r:id") or sheet_el.get("id", "")

                if rid in rels:
                    target = rels[rid]["target"]
                    # Strip leading slash from absolute paths (e.g. /xl/worksheets/sheet1.xml)
                    target = target.lstrip("/")
                    if not target.startswith("xl/"):
                        target = f"xl/{target}"
                    sheet_map[name] = target
                elif name:
                    # Try to find by index
                    idx = len(sheet_map) + 1
                    candidate = f"xl/worksheets/sheet{idx}.xml"
                    if candidate in zip_file.namelist():
                        sheet_map[name] = candidate
        except Exception as e:
            self.log.warning(f"Failed to get sheet map: {e}")

        # Fallback: scan for sheet files
        if not sheet_map:
            for name in zip_file.namelist():
                if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                    sheet_idx = re.search(r"sheet(\d+)\.xml", name).group(1)
                    sheet_map[f"Sheet{sheet_idx}"] = name

        return sheet_map

    def _extract_from_sheet(
        self, zip_file: zipfile.ZipFile, sheet_name: str, sheet_file: str
    ) -> list[DataConnection]:
        """Extract formula-based connections from a single sheet."""
        results = []

        if sheet_file not in zip_file.namelist():
            return results

        try:
            data = zip_file.read(sheet_file)
            root = etree.fromstring(data)
        except Exception as e:
            self.log.warning(f"Failed to parse {sheet_file}: {e}")
            return results

        # Find all cells with formulas
        # Cells are in <row><c><f>formula</f></c></row>
        ns_map = {
            "ss": NS,
        }

        # Find formula elements
        formula_els = root.findall(f".//{{{NS}}}f")
        if not formula_els:
            formula_els = root.findall(".//f")
        if not formula_els:
            formula_els = root.findall(".//{*}f")

        for f_el in formula_els:
            try:
                formula = f_el.text or ""
                if not formula:
                    continue

                # Get parent cell element for address
                parent = f_el.getparent()
                cell_ref = parent.get("r", "") if parent is not None else ""
                location = f"{sheet_name}!{cell_ref}" if cell_ref else sheet_name

                # Prefix formula with = for pattern matching
                formula_str = "=" + formula if not formula.startswith("=") else formula

                found = self._extract_from_formula(formula_str, location)
                results.extend(found)
            except Exception as e:
                self.log.debug(f"Error processing formula cell: {e}")

        return results

    def _extract_from_formula(self, formula: str, location: str) -> list[DataConnection]:
        """Extract connections from a single formula string."""
        results = []

        # Check for WEBSERVICE
        ws_match = WEBSERVICE_PATTERN.search(formula)
        if ws_match:
            url = ws_match.group(1).strip('"\'')
            conn = DataConnection(
                id=DataConnection.make_id("formula", url, location),
                category="web",
                sub_type="webservice",
                source=url[:100],
                raw_connection=url,
                location=location,
                query_text=formula,
                confidence=0.95,
            )
            results.append(conn)

        # Check for RTD
        if RTD_PATTERN.search(formula):
            conn = DataConnection(
                id=DataConnection.make_id("formula", "RTD", location),
                category="formula",
                sub_type="rtd",
                source="RTD (Real-Time Data)",
                raw_connection=formula[:200],
                location=location,
                query_text=formula,
                confidence=0.8,
            )
            results.append(conn)

        # Check for external workbook references
        ext_match = EXTERNAL_WB_PATTERN.search(formula)
        unc_match = UNC_PATTERN.search(formula)
        local_match = LOCAL_PATH_PATTERN.search(formula)

        if ext_match or unc_match or local_match:
            parsed = parse_formula(formula)
            if parsed:
                workbook = parsed.get("workbook_name") or parsed.get("workbook_path", "")
                path = parsed.get("workbook_path", "")
                source = workbook or path or formula[:60]

                # Determine if it's a file or network path
                if path.startswith("\\\\"):
                    sub_type = "unc_path"
                    category = "file"
                elif path and re.match(r"[A-Za-z]:\\", path):
                    sub_type = "local_file"
                    category = "file"
                else:
                    sub_type = "external_workbook"
                    category = "file"

                conn = DataConnection(
                    id=DataConnection.make_id(category, workbook or path, location),
                    category=category,
                    sub_type=sub_type,
                    source=source,
                    raw_connection=path or workbook or formula[:200],
                    location=location,
                    query_text=formula,
                    metadata=parsed,
                    confidence=0.95,
                )
                results.append(conn)

        return results
