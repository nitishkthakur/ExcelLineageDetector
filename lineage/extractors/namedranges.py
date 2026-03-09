"""Extractor for named ranges with external references."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection
from lineage.parsers.formula_parser import parse as parse_formula


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Pattern to detect external workbook references in named range formulas
EXTERNAL_WB_PATTERN = re.compile(
    r"'?(?:[A-Za-z]:\\[^'\[]*|\\\\[^'\[]*)?(?:\[([^\]]+\.(?:xlsx?|xlsm|xlsb|csv))\])",
    re.IGNORECASE,
)


class NamedRangesExtractor(BaseExtractor):
    """Extracts external references from named ranges (defined names)."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            connections.extend(self._extract_from_workbook(zip_file))
        except Exception as e:
            self.log.error(f"NamedRangesExtractor failed: {e}", exc_info=True)
        return connections

    def _extract_from_workbook(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract named ranges from workbook.xml."""
        results = []

        root = self._read_xml(zip_file, "xl/workbook.xml")
        if root is None:
            return results

        # Search all definedName elements anywhere in the document
        # (handles duplicate definedNames elements from XML injection)
        defined_names = root.findall(f".//{{{NS}}}definedName")
        if not defined_names:
            defined_names = root.findall(".//{*}definedName")
        if not defined_names:
            defined_names = root.findall(".//definedName")

        for dn_el in defined_names:
            try:
                name = dn_el.get("name", "")
                formula = dn_el.text or ""
                hidden = dn_el.get("hidden", "0")
                local_sheet_id = dn_el.get("localSheetId", "")

                if not formula:
                    continue

                # Check if formula contains external workbook reference
                if not EXTERNAL_WB_PATTERN.search(formula):
                    continue

                parsed = parse_formula(formula)
                if not parsed:
                    continue

                workbook_name = parsed.get("workbook_name", "")
                workbook_path = parsed.get("workbook_path", "")
                sheet = parsed.get("sheet", "")

                source = workbook_name or workbook_path or formula[:60]
                raw_connection = workbook_path or workbook_name or formula

                location = "workbook.xml:definedNames"
                if local_sheet_id:
                    location = f"Sheet{local_sheet_id}:definedNames"

                conn = DataConnection(
                    id=DataConnection.make_id("file", raw_connection, f"{location}:{name}"),
                    category="file",
                    sub_type="named_range_external",
                    source=source,
                    raw_connection=raw_connection,
                    location=location,
                    query_text=formula,
                    metadata={
                        "defined_name": name,
                        "formula": formula,
                        "hidden": hidden,
                        "local_sheet_id": local_sheet_id,
                        "parsed": parsed,
                    },
                    confidence=0.9,
                )
                results.append(conn)

            except Exception as e:
                self.log.debug(f"Error processing defined name: {e}")

        return results
