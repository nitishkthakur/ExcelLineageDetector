"""Extra scanners covering input-detection gaps not handled by the main pipeline.

Gap coverage implemented here
──────────────────────────────
CRITICAL
  ✓ Transitive missing files   — report design: MissingFile.transitive_unknown flag
  ✓ INDIRECT / dynamic refs    — scan_dynamic_indirect_refs()
  ✓ XLSB files                 — detect_xlsb_files() + optional pyxlsb vector scan

EASY
  ✓ Chart series external refs — scan_chart_external_refs()
  ✓ Data validation sources    — scan_data_validation_refs()
  ✓ Phantom / stale links      — detect_phantom_links()
  ✓ RTD function calls         — scan_rtd_refs()
  ✓ Copy-paste context         — get_vector_context() (header + row label)
  ✓ Scenario Manager entries   — scan_scenarios()

Remaining (hard / out-of-scope for static analysis)
  ✗ VBA file I/O beyond ADODB  — requires VBA binary decompilation
  ✗ Power Pivot / Data Model   — xl/model/item.data is a complex binary blob
  ✗ Finance terminal detail    — BDP/BDH arg extraction needs a separate pass
  ✗ DDE links                  — deprecated, near-zero prevalence
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from lxml import etree

from lineage.hardcoded_scanner import _get_sheet_map, _CELL_RE
from lineage.tracing.formula_tracer import _REF_RE, _get_link_map, _resolve_file

from .models import (
    DynamicRef,
    FormulaRef,
    MissingFile,
    PhantomLink,
    RTDRef,
    ScenarioEntry,
    XlsbWarning,
)


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _stream_formulas_containing(data: bytes, keyword: str) -> list[tuple[str, str]]:
    """Stream sheet XML; return (cell_ref, formula) for cells whose formula
    contains *keyword* (case-insensitive)."""
    kw_upper = keyword.upper()
    results: list[tuple[str, str]] = []
    in_cell = False
    cell_ref = ""
    has_formula = False
    formula_text = ""

    try:
        ctx = etree.iterparse(io.BytesIO(data), events=("start", "end"),
                              recover=True, no_network=True)
        for event, elem in ctx:
            ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if event == "start":
                if ltag == "c":
                    in_cell = True
                    has_formula = False
                    formula_text = ""
                    cell_ref = elem.get("r", "")
                elif ltag == "f" and in_cell:
                    has_formula = True
            else:
                if ltag == "f" and in_cell and has_formula:
                    formula_text = elem.text or ""
                elif ltag == "c":
                    if in_cell and has_formula and formula_text and cell_ref:
                        if kw_upper in formula_text.upper():
                            results.append((cell_ref, formula_text))
                    in_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()
        del ctx
    except Exception:
        pass
    return results


def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Load xl/sharedStrings.xml into a plain list of strings."""
    ss: list[str] = []
    if "xl/sharedStrings.xml" not in set(zf.namelist()):
        return ss
    try:
        root = etree.fromstring(zf.read("xl/sharedStrings.xml"))
        for si in root.iter():
            ltag = si.tag.split("}")[-1] if "}" in si.tag else si.tag
            if ltag == "si":
                parts: list[str] = []
                for child in si.iter():
                    cltag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    if cltag == "t" and child.text:
                        parts.append(child.text)
                ss.append("".join(parts))
    except Exception:
        pass
    return ss


# ---------------------------------------------------------------------------
# 1. Dynamic INDIRECT — CRITICAL
# ---------------------------------------------------------------------------

_INDIRECT_RE = re.compile(r"\bINDIRECT\s*\(", re.IGNORECASE)


def scan_dynamic_indirect_refs(
    path: Path,
    model_sheets: list[str] | None = None,
) -> list[DynamicRef]:
    """Find INDIRECT() formulas where the external file cannot be statically resolved.

    Static INDIRECT like INDIRECT("'[file.xlsx]Sheet1'!A1") are already fully
    handled by formula_tracer (because _REF_RE matches the '[' inside the string
    literal).  This scanner finds the remainder: formulas where the filename is
    assembled at runtime, e.g.:
        =INDIRECT("'["&B1&"]Sheet1'!A1")   — dynamic external (flagged)
        =INDIRECT(A1)                        — fully dynamic (flagged if '[' nearby)
    """
    results: list[DynamicRef] = []
    sheet_set = set(model_sheets) if model_sheets else None

    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            names = set(zf.namelist())
            for sheet_name, sheet_path in sheet_map.items():
                if sheet_set and sheet_name not in sheet_set:
                    continue
                if sheet_path not in names:
                    continue
                try:
                    data = zf.read(sheet_path)
                    formulas = _stream_formulas_containing(data, "INDIRECT")
                    for cell_ref, formula in formulas:
                        # If _REF_RE already finds a file, formula_tracer handles it
                        if _REF_RE.search(formula):
                            continue
                        # Only flag if there are signs of an external reference
                        if not _INDIRECT_RE.search(formula):
                            continue
                        if "[" in formula:
                            note = (
                                "INDIRECT formula contains '[' but no resolvable file "
                                "reference — filename likely assembled from cell values "
                                "at runtime (e.g. INDIRECT(\"'[\"&A1&\"]Sheet!A1\")). "
                                "External dependency invisible to static analysis."
                            )
                        else:
                            # INDIRECT(cell_ref) — cell might hold an external path
                            note = (
                                "INDIRECT formula with cell-reference argument — if the "
                                "referenced cell contains an external workbook path, "
                                "this is an invisible dependency."
                            )
                        results.append(DynamicRef(
                            source_file=path.name,
                            source_sheet=sheet_name,
                            source_cell=cell_ref,
                            formula=formula[:200],
                            note=note,
                        ))
                except Exception:
                    pass
    except Exception:
        pass
    return results


# ---------------------------------------------------------------------------
# 2. Chart series external references — EASY
# ---------------------------------------------------------------------------

def scan_chart_external_refs(
    path: Path,
    search_dirs: list[Path] | None = None,
) -> list[FormulaRef]:
    """Scan xl/charts/*.xml for external workbook references in series data.

    Charts store data-series cell ranges like:
        <c:f>'[actuals_2023.xlsx]Data'!$B$2:$B$13</c:f>
    These are external dependencies but NOT in any worksheet formula.
    Returns FormulaRef objects with ref_origin="chart" so they integrate
    naturally into the rest of the pipeline.
    """
    results: list[FormulaRef] = []
    search = search_dirs or [path.parent]
    seen: set[tuple] = set()

    try:
        with zipfile.ZipFile(path) as zf:
            link_map = _get_link_map(zf)
            names = set(zf.namelist())

            chart_paths = [
                n for n in names
                if (n.startswith("xl/charts/") or n.startswith("xl/chartsheets/"))
                and n.endswith(".xml")
                and "/_rels/" not in n
            ]

            for chart_path in chart_paths:
                chart_name = chart_path.rsplit("/", 1)[-1]
                try:
                    data = zf.read(chart_path)
                    root = etree.fromstring(data)

                    for elem in root.iter():
                        ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                        if ltag != "f" or not elem.text or "[" not in elem.text:
                            continue
                        formula = elem.text

                        for m in _REF_RE.finditer(formula):
                            file_id = m.group(1)
                            sheet = m.group(2).strip()
                            cell_range = m.group(3).replace("$", "")

                            if file_id.isdigit():
                                resolved_name = link_map.get(file_id)
                                filename = resolved_name[0] if resolved_name else f"[ExternalLink{file_id}]"
                            else:
                                filename = file_id

                            key = (filename.lower(), sheet.lower(), cell_range.upper())
                            if key in seen:
                                continue
                            seen.add(key)

                            resolved, found = _resolve_file(filename, search)
                            results.append(FormulaRef(
                                level=1,
                                source_file=path.name,
                                source_sheet=f"(chart:{chart_name})",
                                source_cell="",
                                target_file=filename,
                                target_sheet=sheet,
                                target_range=cell_range,
                                file_found=found,
                                resolved_path=str(resolved),
                                ref_origin="chart",
                            ))
                except Exception:
                    pass
    except Exception:
        pass
    return results


# ---------------------------------------------------------------------------
# 3. Data validation external sources — EASY
# ---------------------------------------------------------------------------

def scan_data_validation_refs(
    path: Path,
    search_dirs: list[Path] | None = None,
) -> list[FormulaRef]:
    """Scan <dataValidation> elements for external workbook list sources.

    A dropdown list like:
        Data Validation → List → Source: '[master.xlsx]Lists'!$A:$A
    is stored in <formula1> inside <dataValidation>.  If the source file is
    missing, all dropdown cells show invalid or empty options.
    Returns FormulaRef with ref_origin="data_validation".
    """
    results: list[FormulaRef] = []
    search = search_dirs or [path.parent]
    seen: set[tuple] = set()

    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            link_map = _get_link_map(zf)
            names = set(zf.namelist())

            for sheet_name, sheet_path in sheet_map.items():
                if sheet_path not in names:
                    continue
                try:
                    data = zf.read(sheet_path)
                    # Use iterparse to handle large sheets without loading the full tree
                    ctx = etree.iterparse(io.BytesIO(data), events=("start", "end"),
                                         recover=True, no_network=True)
                    in_dv = False
                    sqref = ""
                    for event, elem in ctx:
                        ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                        if event == "start":
                            if ltag == "dataValidation":
                                in_dv = True
                                sqref = elem.get("sqref", "")
                        else:
                            if ltag in ("formula1", "formula2") and in_dv:
                                formula = elem.text or ""
                                if "[" in formula:
                                    for m in _REF_RE.finditer(formula):
                                        file_id = m.group(1)
                                        sheet = m.group(2).strip()
                                        cell_range = m.group(3).replace("$", "")
                                        if file_id.isdigit():
                                            rn = link_map.get(file_id)
                                            filename = rn[0] if rn else f"[ExternalLink{file_id}]"
                                        else:
                                            filename = file_id
                                        key = (filename.lower(), sheet.lower(), cell_range.upper())
                                        if key in seen:
                                            continue
                                        seen.add(key)
                                        resolved, found = _resolve_file(filename, search)
                                        results.append(FormulaRef(
                                            level=1,
                                            source_file=path.name,
                                            source_sheet=sheet_name,
                                            source_cell=f"(validation:{sqref})",
                                            target_file=filename,
                                            target_sheet=sheet,
                                            target_range=cell_range,
                                            file_found=found,
                                            resolved_path=str(resolved),
                                            ref_origin="data_validation",
                                        ))
                            elif ltag == "dataValidation":
                                in_dv = False
                            elif ltag == "row":
                                elem.clear()
                    del ctx
                except Exception:
                    pass
    except Exception:
        pass
    return results


# ---------------------------------------------------------------------------
# 4. RTD (Real-Time Data) function calls — EASY
# ---------------------------------------------------------------------------

_RTD_PROGID_RE = re.compile(r'\bRTD\s*\(\s*"([^"]*)"', re.IGNORECASE)


def scan_rtd_refs(
    path: Path,
    model_sheets: list[str] | None = None,
) -> list[RTDRef]:
    """Detect =RTD() function calls in formulas.

    RTD delivers live data via a COM server (Bloomberg, Reuters Eikon, etc.).
    These are NOT files or ODBC connections — the data feed must be running.
    Example: =RTD("bloomberg.rtd",,"AAPL US Equity","LAST_PRICE")
    """
    results: list[RTDRef] = []
    sheet_set = set(model_sheets) if model_sheets else None
    seen: set[tuple] = set()

    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            names = set(zf.namelist())
            for sheet_name, sheet_path in sheet_map.items():
                if sheet_set and sheet_name not in sheet_set:
                    continue
                if sheet_path not in names:
                    continue
                try:
                    data = zf.read(sheet_path)
                    formulas = _stream_formulas_containing(data, "RTD")
                    for cell_ref, formula in formulas:
                        m = _RTD_PROGID_RE.search(formula)
                        prog_id = m.group(1) if m else ""
                        key = (sheet_name, cell_ref)
                        if key in seen:
                            continue
                        seen.add(key)
                        results.append(RTDRef(
                            source_file=path.name,
                            source_sheet=sheet_name,
                            source_cell=cell_ref,
                            prog_id=prog_id,
                            formula=formula[:200],
                        ))
                except Exception:
                    pass
    except Exception:
        pass
    return results


# ---------------------------------------------------------------------------
# 5. Phantom / stale external links — EASY
# ---------------------------------------------------------------------------

def detect_phantom_links(
    path: Path,
    formula_ref_files: set[str],   # lower-cased filenames actually referenced
) -> list[PhantomLink]:
    """Find xl/externalLinks/ entries that no formula actually uses.

    Excel retains stale links even after all referencing formulas are deleted.
    They appear as dependencies but aren't needed and should be broken.
    """
    phantoms: list[PhantomLink] = []
    try:
        with zipfile.ZipFile(path) as zf:
            link_map = _get_link_map(zf)
            for _idx, (filename, _full_path) in link_map.items():
                if filename.lower() not in formula_ref_files:
                    phantoms.append(PhantomLink(
                        source_file=path.name,
                        stale_filename=filename,
                    ))
    except Exception:
        pass
    return phantoms


# ---------------------------------------------------------------------------
# 6. Unmatched vector context — EASY
# ---------------------------------------------------------------------------

def get_vector_context(
    path: Path,
    sheet_name: str,
    start_cell: str,
) -> tuple[str, str]:
    """Return (column_header, row_label) for a vector's start cell.

    column_header: value of row-1 in the vector's column (e.g. B1 for vector B5:B17)
    row_label:     value of column-A in the vector's start row (e.g. A5)

    These give analysts context about what the hardcoded data represents,
    even when no upstream source was found.
    """
    m = _CELL_RE.match(start_cell.replace("$", ""))
    if not m:
        return ("", "")

    col_str = m.group(1).upper()
    start_row = int(m.group(2))

    header_ref = f"{col_str}1"
    label_ref = f"A{start_row}"
    targets = {header_ref, label_ref}
    found: dict[str, str] = {}

    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            if sheet_name not in sheet_map:
                return ("", "")
            sheet_path = sheet_map[sheet_name]
            if sheet_path not in set(zf.namelist()):
                return ("", "")

            ss = _load_shared_strings(zf)
            data = zf.read(sheet_path)

            in_cell = False
            cell_ref = ""
            cell_type = "n"
            pending_v = ""

            ctx = etree.iterparse(io.BytesIO(data), events=("start", "end"),
                                  recover=True, no_network=True)
            for event, elem in ctx:
                if len(found) == len(targets):
                    break
                ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if event == "start":
                    if ltag == "c":
                        in_cell = True
                        pending_v = ""
                        raw_ref = elem.get("r", "").replace("$", "").upper()
                        # strip row from ref, rebuild cleanly
                        mm = _CELL_RE.match(raw_ref)
                        cell_ref = f"{mm.group(1)}{mm.group(2)}" if mm else raw_ref
                        cell_type = elem.get("t", "n")
                else:
                    if ltag == "v" and in_cell:
                        pending_v = elem.text or ""
                    elif ltag == "c":
                        if in_cell and cell_ref in targets and pending_v:
                            if cell_type == "s":
                                try:
                                    found[cell_ref] = ss[int(pending_v)]
                                except (IndexError, ValueError):
                                    found[cell_ref] = pending_v
                            else:
                                found[cell_ref] = pending_v
                        in_cell = False
                        elem.clear()
                    elif ltag == "row":
                        elem.clear()
            del ctx
    except Exception:
        pass

    return (found.get(header_ref, ""), found.get(label_ref, ""))


# ---------------------------------------------------------------------------
# 7. XLSB file detection — CRITICAL
# ---------------------------------------------------------------------------

def detect_xlsb_files(
    inputs_dir: Path,
    formula_refs: list[FormulaRef],
) -> list[XlsbWarning]:
    """Find .xlsb files — binary Excel format that cannot be ZIP/XML analysed.

    Checks two sources:
    1. .xlsb files found inside inputs_dir (can't be scanned for formulas or vectors)
    2. .xlsb files referenced by formula tracing (model depends on them)

    If pyxlsb is installed, basic vector extraction is possible for upstream
    matching; formula scanning is still not supported.
    """
    try:
        import pyxlsb  # noqa: F401
        pyxlsb_ok = True
    except ImportError:
        pyxlsb_ok = False

    warnings: list[XlsbWarning] = []
    seen: set[str] = set()

    # Inputs dir scan
    for p in inputs_dir.rglob("*.xlsb"):
        key = p.name.lower()
        if key not in seen:
            seen.add(key)
            warnings.append(XlsbWarning(
                filename=p.name,
                path=str(p),
                source="inputs_dir",
                pyxlsb_available=pyxlsb_ok,
            ))

    # Formula refs targeting xlsb
    for ref in formula_refs:
        if ref.target_file.lower().endswith(".xlsb"):
            key = ref.target_file.lower()
            if key not in seen:
                seen.add(key)
                warnings.append(XlsbWarning(
                    filename=ref.target_file,
                    path=ref.resolved_path if ref.file_found else ref.target_file,
                    source=ref.source_file,
                    pyxlsb_available=pyxlsb_ok,
                ))

    return warnings


def scan_xlsb_vectors(path: Path, min_len: int = 3) -> list:
    """Extract numeric vectors from an .xlsb file using pyxlsb (if available).

    Returns list of TracingVector, same format as scan_upstream_file().
    Returns [] silently if pyxlsb is not installed.
    """
    try:
        import pyxlsb
    except ImportError:
        return []

    cells_by_sheet: dict[str, list[tuple[int, int, float]]] = {}
    try:
        with pyxlsb.open_workbook(str(path)) as wb:
            for sheet_name in wb.sheets:
                cells: list[tuple[int, int, float]] = []
                try:
                    with wb.get_sheet(sheet_name) as ws:
                        for row_data in ws.rows():
                            for cell in row_data:
                                if cell.v is not None:
                                    try:
                                        cells.append((cell.r, cell.c, float(cell.v)))
                                    except (TypeError, ValueError):
                                        pass
                except Exception:
                    pass
                if cells:
                    cells_by_sheet[sheet_name] = cells
    except Exception:
        return []

    from lineage.tracing.scanner import _cells_to_vectors
    vectors = []
    for sheet_name, cells in cells_by_sheet.items():
        vectors.extend(_cells_to_vectors(cells, path.name, sheet_name, min_len))
    return vectors


# ---------------------------------------------------------------------------
# 8. Scenario Manager entries — EASY
# ---------------------------------------------------------------------------

def scan_scenarios(path: Path) -> list[ScenarioEntry]:
    """Extract named scenarios from Excel's Scenario Manager.

    Scenarios (Bull/Base/Bear case, etc.) store sets of input values in
    <scenarios> elements within worksheet XML.  They are inputs that do not
    appear in any formula and are completely invisible to formula tracing.
    """
    results: list[ScenarioEntry] = []
    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            names = set(zf.namelist())
            for sheet_name, sheet_path in sheet_map.items():
                if sheet_path not in names:
                    continue
                try:
                    data = zf.read(sheet_path)
                    ctx = etree.iterparse(io.BytesIO(data), events=("start", "end"),
                                         recover=True, no_network=True)
                    in_scenarios = False
                    current_scenario: ScenarioEntry | None = None

                    for event, elem in ctx:
                        ltag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                        if event == "start":
                            if ltag == "scenarios":
                                in_scenarios = True
                            elif ltag == "scenario" and in_scenarios:
                                current_scenario = ScenarioEntry(
                                    source_file=path.name,
                                    sheet_name=sheet_name,
                                    scenario_name=elem.get("name", ""),
                                    input_cells=[],
                                )
                            elif ltag == "inputCells" and current_scenario is not None:
                                r = elem.get("r", "")
                                v = elem.get("val", "")
                                if r:
                                    current_scenario.input_cells.append((r, v))
                        else:
                            if ltag == "scenario" and current_scenario is not None:
                                if current_scenario.scenario_name:
                                    results.append(current_scenario)
                                current_scenario = None
                            elif ltag == "scenarios":
                                in_scenarios = False
                            elif ltag == "row":
                                elem.clear()
                    del ctx
                except Exception:
                    pass
    except Exception:
        pass
    return results
