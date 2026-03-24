"""Data models for RawSourcesDetection."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class SourceNode:
    """A file in the dependency graph (model or upstream)."""
    filename: str
    path: str               # full disk path if found; expected name if missing
    level: int              # 0 = model, 1 = direct upstream, 2 = upstream's upstream, …
    found_on_disk: bool
    sheets_referenced: list[str] = field(default_factory=list)


@dataclass
class FormulaRef:
    """One external formula reference found during tracing."""
    level: int              # at which level this ref was found
    source_file: str        # file containing the formula
    source_sheet: str
    source_cell: str
    target_file: str        # filename referenced in the formula
    target_sheet: str
    target_range: str
    file_found: bool        # whether target_file was found in inputs/
    resolved_path: str      # full disk path if found
    ref_origin: str = "formula"   # "formula" | "chart" | "data_validation"


@dataclass
class MissingFile:
    """A file referenced in formulas but not found in inputs/."""
    filename: str
    level: int              # tracing level at which it was first discovered
    referenced_by: str      # source file that references it
    sheets_needed: list[str] = field(default_factory=list)
    cells_referencing: list[str] = field(default_factory=list)  # "Sheet!Cell" strings
    # Transitive: we cannot know what THIS missing file itself depends on.
    # All its downstream dependencies are invisible until the file is supplied.
    transitive_unknown: bool = True


@dataclass
class MatchedVector:
    """A hardcoded model vector matched to an upstream source."""
    model_sheet: str
    model_range: str
    model_length: int
    model_sample: list[float]   # first 5 values
    match_type: str             # "exact" | "exact_subsequence" | "approximate"
    similarity: float           # 1.0 for exact matches
    upstream_file: str
    upstream_sheet: str
    upstream_range: str
    upstream_sample: list[float]


@dataclass
class UnmatchedVector:
    """A hardcoded model vector with no source match found."""
    model_sheet: str
    model_range: str
    model_length: int
    model_sample: list[float]
    column_header: str = ""   # value of row-1 cell above the vector (context)
    row_label: str = ""       # value of column-A cell beside the vector (context)


@dataclass
class RawSource:
    """A raw data source connection (ODBC, OLE DB, Power Query, Excel link, etc.)."""
    source_file: str    # which file this was detected in
    category: str       # database | file | powerquery | formula | vba | …
    sub_type: str       # odbc | oledb | sql_server | xlsx | csv | …
    connection: str     # raw connection string / path / URL
    location: str       # where in the file it was found (e.g. "Sheet1!A1")


# ---------------------------------------------------------------------------
# New models from extra scanners
# ---------------------------------------------------------------------------

@dataclass
class DynamicRef:
    """An INDIRECT() formula whose external file target cannot be statically resolved.

    Static INDIRECT like INDIRECT("'[file.xlsx]Sheet1'!A1") are already caught
    by formula_tracer.  These are the dangerous ones — e.g.:
        =INDIRECT("'["&A1&"]Sheet1'!A1")
    where the filename is assembled at runtime from cell values.
    """
    source_file: str
    source_sheet: str
    source_cell: str
    formula: str        # truncated to 200 chars
    note: str           # human-readable explanation


@dataclass
class RTDRef:
    """A Real-Time Data (RTD) function call — live COM data feed.

    Example: =RTD("bloomberg.rtd",,"AAPL US Equity","LAST_PRICE")
    These are NOT files or ODBC connections — they require a running COM server
    (Bloomberg Terminal, Reuters Eikon, etc.) to return data.
    """
    source_file: str
    source_sheet: str
    source_cell: str
    prog_id: str        # e.g. "bloomberg.rtd", "ek.rtd"
    formula: str


@dataclass
class PhantomLink:
    """A stale external link registered in xl/externalLinks/ but not used by any formula.

    These inflate the apparent dependency list. They should be removed via
    Excel → Data → Edit Links → Break Link.
    """
    source_file: str
    stale_filename: str     # the file that used to be referenced


@dataclass
class XlsbWarning:
    """An .xlsb (binary Excel) file that cannot be fully analysed.

    xlsb is a ZIP-incompatible binary format. Formula scanning and vector
    matching require ZIP/XML access and are not possible without pyxlsb.
    """
    filename: str
    path: str           # disk path if found, else just the filename
    source: str         # "inputs_dir" | source file that references it
    pyxlsb_available: bool = False


@dataclass
class ScenarioEntry:
    """One named scenario from Excel's Scenario Manager.

    Scenarios represent stored sets of input values (Bull / Base / Bear cases).
    They are inputs but appear nowhere in formulas — only in the Scenario Manager UI.
    """
    source_file: str
    sheet_name: str
    scenario_name: str
    input_cells: list[tuple[str, str]] = field(default_factory=list)  # (cell_ref, value)


@dataclass
class DetectionResult:
    """Complete output from the RawSourcesDetection pipeline."""
    model_file: str
    source_nodes: list[SourceNode]            # all files in the dependency tree
    formula_refs: list[FormulaRef]            # formula + chart + data-validation refs
    missing_files: list[MissingFile]          # files referenced but not on disk
    matched_vectors: list[MatchedVector]      # hardcoded vectors with source found
    unmatched_vectors: list[UnmatchedVector]  # hardcoded vectors with no source (with context)
    raw_sources: list[RawSource]              # ODBC/OLE DB/PQ connections
    # Extra scanner results
    dynamic_refs: list[DynamicRef] = field(default_factory=list)
    rtd_refs: list[RTDRef] = field(default_factory=list)
    phantom_links: list[PhantomLink] = field(default_factory=list)
    xlsb_warnings: list[XlsbWarning] = field(default_factory=list)
    scenarios: list[ScenarioEntry] = field(default_factory=list)
