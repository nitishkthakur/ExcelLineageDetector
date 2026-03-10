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
# Also handles HTTPS/SharePoint prefix: 'https://company.sharepoint.com/...[wb.xlsx]'
EXTERNAL_WB_PATTERN = re.compile(
    r"'?(?:[A-Za-z]:\\[^'\[]*|\\\\[^'\[]*|https?://[^'\[]+)?(?:\[([^\]]+\.(?:xlsx?|xlsm|xlsb|csv))\])",
    re.IGNORECASE,
)

# UNC path in formula
UNC_PATTERN = re.compile(r"'(\\\\[^'\[]+)\[", re.IGNORECASE)

# Local drive path in formula
LOCAL_PATH_PATTERN = re.compile(r"'([A-Za-z]:\\[^'\[]+)\[", re.IGNORECASE)

# HTTPS/SharePoint/OneDrive path in formula:
#   'https://company.sharepoint.com/sites/dept/[budget.xlsx]Sheet1'!A1
HTTPS_PATH_PATTERN = re.compile(
    r"'(https?://[^']+)\[([^\]]+\.(?:xlsx?|xlsm|xlsb|csv))\]([^'!]*)(?:')?!",
    re.IGNORECASE,
)

# Numeric external link index: =[1]Sheet1!A1 or =[2]Data!B5
# The index corresponds to xl/externalLinks/externalLink{n}.xml
EXTERNAL_INDEX_PATTERN = re.compile(
    r"(?:^|[=(+\-*,\s])\[(\d+)\]([^![]+)!",
    re.IGNORECASE,
)

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

# INDIRECT function - dynamic reference (path not statically determinable)
INDIRECT_PATTERN = re.compile(r'(?i)\bINDIRECT\s*\(', re.IGNORECASE)

# Bloomberg Terminal live data formulas
BLOOMBERG_PATTERN = re.compile(
    r'(?i)\b(BDP|BDH|BDS|BQL|BSRCH|BLPGET|BQ)\s*\(\s*"([^"]*)"(?:\s*,\s*"([^"]*)")?',
    re.IGNORECASE,
)
# Reuters/Refinitiv Eikon formulas
REUTERS_PATTERN = re.compile(
    r'(?i)\b(RHistory|RData|RKFGET|TR\.)\s*\(',
    re.IGNORECASE,
)
# FactSet formulas
FACTSET_PATTERN = re.compile(
    r'(?i)\b(FDS|FQL|FDSC|FEW|FDSQ)\s*\(',
    re.IGNORECASE,
)
# Capital IQ / S&P formulas
CAPITALIQ_PATTERN = re.compile(
    r'(?i)\b(CIQCONTENT|CIQ|SPCIQDATA|CIQ_CONTENT|CIQ_FORMULAFIELD|SCS)\s*\(',
    re.IGNORECASE,
)
# SNL Financial formulas
SNL_PATTERN = re.compile(r'(?i)\bSNL[DC]?\s*\(', re.IGNORECASE)
# Wind Information (Chinese financial terminal)
WIND_PATTERN = re.compile(r'(?i)\bW(?:SD|SS|Q|QS|SP)\s*\(', re.IGNORECASE)


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

        # Scan data validation formulas (can contain external list references)
        try:
            results.extend(self._extract_from_data_validation(root, sheet_name))
        except Exception as e:
            self.log.debug(f"Error scanning data validation in {sheet_name}: {e}")

        # Scan conditional formatting formulas
        try:
            results.extend(self._extract_from_conditional_formatting(root, sheet_name))
        except Exception as e:
            self.log.debug(f"Error scanning conditional formatting in {sheet_name}: {e}")

        return results

    def _extract_from_data_validation(self, root, sheet_name: str) -> list[DataConnection]:
        """Scan dataValidation formula1/formula2 elements for external references."""
        results = []
        dv_els = (
            root.findall(f".//{{{NS}}}dataValidation")
            or root.findall(".//dataValidation")
            or root.findall(".//{*}dataValidation")
        )
        for dv_el in dv_els:
            sqref = dv_el.get("sqref", "")
            location = (
                f"{sheet_name}!dataValidation:{sqref}"
                if sqref else f"{sheet_name}:dataValidation"
            )
            for tag_suffix in ["formula1", "formula2"]:
                f_el = (
                    dv_el.find(f"{{{NS}}}{tag_suffix}")
                    or dv_el.find(f"{{*}}{tag_suffix}")
                    or dv_el.find(tag_suffix)
                )
                if f_el is None:
                    continue
                formula = (f_el.text or "").strip()
                if not formula:
                    continue
                formula_str = "=" + formula if not formula.startswith("=") else formula
                results.extend(self._extract_from_formula(formula_str, location))
        return results

    def _extract_from_conditional_formatting(self, root, sheet_name: str) -> list[DataConnection]:
        """Scan conditionalFormatting formula elements for external references."""
        results = []
        cf_rules = (
            root.findall(f".//{{{NS}}}cfRule")
            or root.findall(".//cfRule")
            or root.findall(".//{*}cfRule")
        )
        for cf_el in cf_rules:
            location = f"{sheet_name}:conditionalFormatting"
            f_el = (
                cf_el.find(f"{{{NS}}}formula")
                or cf_el.find("{*}formula")
                or cf_el.find("formula")
            )
            if f_el is None:
                continue
            formula = (f_el.text or "").strip()
            if not formula:
                continue
            formula_str = "=" + formula if not formula.startswith("=") else formula
            results.extend(self._extract_from_formula(formula_str, location))
        return results

    def _extract_from_formula(self, formula: str, location: str) -> list[DataConnection]:
        """Extract connections from a single formula string."""
        results = []

        # Check for Bloomberg Terminal formulas
        bb_match = BLOOMBERG_PATTERN.search(formula)
        if bb_match:
            func = bb_match.group(1).upper()
            ticker = bb_match.group(2)
            field = bb_match.group(3) or ""
            source_label = f"{ticker} [{field}]" if field else ticker
            conn = DataConnection(
                id=DataConnection.make_id("formula", f"bloomberg:{ticker}:{field}", location),
                category="formula",
                sub_type="bloomberg",
                source=source_label[:100],
                raw_connection=formula[:200],
                location=location,
                query_text=formula,
                metadata={"function": func, "ticker": ticker, "field": field},
                confidence=0.95,
            )
            results.append(conn)

        # Check for Reuters/Refinitiv Eikon formulas
        if REUTERS_PATTERN.search(formula):
            conn = DataConnection(
                id=DataConnection.make_id("formula", "reuters:" + formula[:40], location),
                category="formula", sub_type="reuters",
                source="Reuters/Refinitiv Eikon",
                raw_connection=formula[:200], location=location, query_text=formula,
                confidence=0.9,
            )
            results.append(conn)

        # Check for FactSet formulas
        if FACTSET_PATTERN.search(formula):
            conn = DataConnection(
                id=DataConnection.make_id("formula", "factset:" + formula[:40], location),
                category="formula", sub_type="factset",
                source="FactSet",
                raw_connection=formula[:200], location=location, query_text=formula,
                confidence=0.9,
            )
            results.append(conn)

        # Check for Capital IQ / S&P formulas
        if CAPITALIQ_PATTERN.search(formula):
            conn = DataConnection(
                id=DataConnection.make_id("formula", "capitaliq:" + formula[:40], location),
                category="formula", sub_type="capitaliq",
                source="Capital IQ / S&P",
                raw_connection=formula[:200], location=location, query_text=formula,
                confidence=0.9,
            )
            results.append(conn)

        # Check for SNL Financial formulas
        if SNL_PATTERN.search(formula):
            conn = DataConnection(
                id=DataConnection.make_id("formula", "snl:" + formula[:40], location),
                category="formula", sub_type="snl",
                source="SNL Financial",
                raw_connection=formula[:200], location=location, query_text=formula,
                confidence=0.9,
            )
            results.append(conn)

        # Check for Wind Information formulas
        if WIND_PATTERN.search(formula):
            conn = DataConnection(
                id=DataConnection.make_id("formula", "wind:" + formula[:40], location),
                category="formula", sub_type="wind",
                source="Wind Information",
                raw_connection=formula[:200], location=location, query_text=formula,
                confidence=0.9,
            )
            results.append(conn)

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

        # Check for HTTPS/SharePoint/OneDrive path in formula:
        # ='https://company.sharepoint.com/sites/dept/[budget.xlsx]Sheet1'!A1
        https_match = HTTPS_PATH_PATTERN.search(formula)
        if https_match:
            url_path = https_match.group(1)
            workbook = https_match.group(2)
            sheet = https_match.group(3).strip("'")
            full_path = url_path + workbook
            pl = url_path.lower()
            if "sharepoint.com" in pl or "/sharepoint/" in pl:
                category, sub_type = "file", "sharepoint_workbook"
            elif "onedrive" in pl or "live.net" in pl or "1drv.ms" in pl:
                category, sub_type = "file", "onedrive_workbook"
            else:
                category, sub_type = "web", "http_workbook"
            conn = DataConnection(
                id=DataConnection.make_id(category, full_path, location),
                category=category,
                sub_type=sub_type,
                source=workbook,
                raw_connection=full_path,
                location=location,
                query_text=formula,
                metadata={
                    "url_path": url_path,
                    "workbook_name": workbook,
                    "sheet": sheet,
                },
                confidence=0.95,
            )
            results.append(conn)

        # Check for numeric external link index references: =[1]Sheet1!A1
        # The index maps to xl/externalLinks/externalLink{n}.xml
        for idx_match in EXTERNAL_INDEX_PATTERN.finditer(formula):
            link_idx = idx_match.group(1)
            sheet_ref = idx_match.group(2).strip()
            ref_str = f"[{link_idx}]{sheet_ref}"
            conn = DataConnection(
                id=DataConnection.make_id("file", ref_str, location),
                category="file",
                sub_type="external_workbook_ref",
                source=f"External Workbook [{link_idx}]",
                raw_connection=ref_str,
                location=location,
                query_text=formula,
                metadata={
                    "link_index": link_idx,
                    "sheet": sheet_ref,
                    "note": f"Full path in xl/externalLinks/externalLink{link_idx}.xml",
                },
                confidence=0.85,
            )
            results.append(conn)

        # Check for INDIRECT() - dynamic references (can't resolve statically)
        if INDIRECT_PATTERN.search(formula):
            # Only flag if it contains string literals that look like external refs
            if re.search(r'INDIRECT\s*\(["\'][^"\']*\[', formula, re.IGNORECASE):
                conn = DataConnection(
                    id=DataConnection.make_id("formula", "INDIRECT:" + formula[:50], location),
                    category="formula",
                    sub_type="indirect_ref",
                    source="INDIRECT (dynamic reference)",
                    raw_connection=formula[:200],
                    location=location,
                    query_text=formula,
                    confidence=0.6,
                )
                results.append(conn)

        # Check for external workbook references (local/UNC paths and bare [file.xlsx])
        ext_match = EXTERNAL_WB_PATTERN.search(formula)
        unc_match = UNC_PATTERN.search(formula)
        local_match = LOCAL_PATH_PATTERN.search(formula)

        # Skip if already handled as HTTPS above
        if (ext_match or unc_match or local_match) and not https_match:
            parsed = parse_formula(formula)
            if parsed:
                workbook = parsed.get("workbook_name") or parsed.get("workbook_path", "")
                path = parsed.get("workbook_path", "")
                source = workbook or path or formula[:60]

                # Determine category/sub_type from path format
                if path.startswith("\\\\"):
                    sub_type = "unc_path"
                    category = "file"
                elif path and re.match(r"[A-Za-z]:\\", path):
                    sub_type = "local_file"
                    category = "file"
                elif path.lower().startswith("https://") or path.lower().startswith("http://"):
                    pl = path.lower()
                    if "sharepoint" in pl:
                        sub_type, category = "sharepoint_workbook", "file"
                    elif "onedrive" in pl or "live.net" in pl:
                        sub_type, category = "onedrive_workbook", "file"
                    else:
                        sub_type, category = "http_workbook", "web"
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
