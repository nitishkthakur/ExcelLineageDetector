"""Recursive formula-based upstream tracing with transitive precedent walking.

Scans Excel files for formulas that reference external workbooks, then
follows those references recursively through multiple levels until the
upstream file is no longer found on disk.

For Level 2+ scans, if a target cell doesn't directly reference an external
file but its formula depends on other cells that do (possibly through
multiple intermediate formulas), the precedent chain is walked transitively
using BFS until the external reference is found.

Uses lxml iterparse for O(n) streaming — same constant-memory approach
as the hardcoded vector scanner.
"""
from __future__ import annotations

import io
import re
import zipfile
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote

from lxml import etree

from lineage.hardcoded_scanner import _col_to_idx, _idx_to_col, _CELL_RE, _get_sheet_map


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Safety limits for precedent walking
_MAX_PRECEDENT_DEPTH = 20        # max BFS depth per starting cell
_MAX_CELLS_VISITED = 10_000      # max cells visited per starting cell
_MAX_RANGE_EXPANSION = 10_000    # max cells from a single range expansion


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class CellFilter:
    """Defines which cells to inspect in a file (for Level 2+ scoping)."""
    # sheet_name → list of (min_row, max_row, min_col, max_col) rectangles
    ranges: dict[str, list[tuple[int, int, int, int]]] = field(default_factory=dict)

    def contains(self, sheet: str, row: int, col: int) -> bool:
        rects = self.ranges.get(sheet)
        if rects is None:
            return False
        return any(r[0] <= row <= r[1] and r[2] <= col <= r[3] for r in rects)

    def has_sheet(self, sheet: str) -> bool:
        return sheet in self.ranges

    @classmethod
    def from_refs(cls, cell_specs: list[tuple[str, str]]) -> CellFilter:
        """Build from list of (sheet_name, cell_range_str)."""
        ranges: dict[str, list[tuple[int, int, int, int]]] = {}
        for sheet, cell_range in cell_specs:
            rect = _parse_range(cell_range)
            ranges.setdefault(sheet, []).append(rect)
        return cls(ranges=ranges)


@dataclass
class ExternalReference:
    """One formula reference from a source cell to an external workbook."""
    level: int
    source_file: str
    source_sheet: str
    source_cell: str
    formula: str
    target_file: str       # resolved filename (e.g. "budget.xlsx")
    target_sheet: str
    target_range: str      # e.g. "A1:B10"
    target_path: str       # full path from formula / rels (for display)
    file_found: bool
    resolved_path: str     # actual disk path if found, else expected path
    # Precedent chain for transitive references (None = direct reference).
    # Each tuple: (sheet_name, cell_ref, formula_snippet).
    precedent_chain: list[tuple[str, str, str]] | None = None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_range(cell_range: str) -> tuple[int, int, int, int]:
    """Parse 'A1:B10' → (min_row, max_row, min_col, max_col)."""
    cell_range = cell_range.replace("$", "")
    parts = cell_range.split(":")

    m1 = _CELL_RE.match(parts[0])
    if not m1:
        return (1, 1048576, 1, 16384)  # full sheet fallback

    col1 = _col_to_idx(m1.group(1))
    row1 = int(m1.group(2))

    if len(parts) == 2:
        m2 = _CELL_RE.match(parts[1])
        if m2:
            col2 = _col_to_idx(m2.group(1))
            row2 = int(m2.group(2))
            return (min(row1, row2), max(row1, row2),
                    min(col1, col2), max(col1, col2))

    return (row1, row1, col1, col1)


def _extract_filename(path: str) -> str:
    """Extract the Excel filename from a path / URL / rels target."""
    path = unquote(path)
    if path.startswith("file:///"):
        path = path[8:]
    elif path.startswith("file://"):
        path = path[7:]
    # Get the basename — handle both / and \ separators
    name = path.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    # Strip any query params or fragments
    name = name.split("?")[0].split("#")[0]
    return name


def _expand_range(cell_range: str) -> list[str]:
    """Expand 'A1:C3' into individual cell references.

    Returns a list of cell refs. Single cells like 'A1' return ['A1'].
    Large ranges are capped at _MAX_RANGE_EXPANSION cells.
    """
    cell_range = cell_range.replace("$", "")
    parts = cell_range.split(":")
    if len(parts) == 1:
        return [parts[0]]

    m1 = _CELL_RE.match(parts[0])
    m2 = _CELL_RE.match(parts[1])
    if not m1 or not m2:
        return [cell_range]

    col1, row1 = _col_to_idx(m1.group(1)), int(m1.group(2))
    col2, row2 = _col_to_idx(m2.group(1)), int(m2.group(2))

    min_r, max_r = min(row1, row2), max(row1, row2)
    min_c, max_c = min(col1, col2), max(col1, col2)

    n_cells = (max_r - min_r + 1) * (max_c - min_c + 1)
    if n_cells > _MAX_RANGE_EXPANSION:
        return []  # too large — skip

    cells: list[str] = []
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            cells.append(f"{_idx_to_col(c)}{r}")
    return cells


# Regex: captures [filename_or_index], sheet_name, and cell_range from formulas.
# Handles:  '[file.xlsx]Sheet1'!A1:B10
#           [1]Sheet1!A1
#           'C:\path\[file.xlsx]Sheet Name'!$A$1
#           'https://sharepoint.com/.../[file.xlsx]Sheet1'!A1
_REF_RE = re.compile(
    r"'?"                                   # optional opening quote
    r"(?:[^'\[\]]*?)"                       # optional path prefix (non-greedy)
    r"\[([^\]]+)\]"                         # group 1: filename or numeric index
    r"([^'!]*)"                             # group 2: sheet name
    r"'?"                                   # optional closing quote
    r"!"                                    # exclamation separator
    r"(\$?[A-Z]{1,3}\$?\d+"                # group 3: cell ref start
    r"(?::\$?[A-Z]{1,3}\$?\d+)?)",         # optional :end
    re.IGNORECASE,
)

# Regex to capture the full path prefix before [filename] in a formula.
_PATH_RE = re.compile(
    r"'([^']*?)\[([^\]]+)\]",
    re.IGNORECASE,
)

# Regex: captures intra-workbook cell references in formulas.
# Matches:  A1, $A$1, A1:B10, $A$1:$B$10
#           Sheet2!A1, Sheet2!A1:B10
#           'Sheet Name'!A1, 'Data Sheet'!A1:B10
# Does NOT match external refs — those are filtered out by overlap with _REF_RE.
_INTRA_REF_RE = re.compile(
    r"(?:"
    r"(?:'([^'\[\]]+)'|([A-Za-z_]\w*))"    # group 1: quoted sheet, group 2: unquoted sheet
    r"!"
    r")?"                                   # sheet prefix is optional
    r"(\$?[A-Z]{1,3}\$?\d+"               # group 3: cell ref (possibly range)
    r"(?::\$?[A-Z]{1,3}\$?\d+)?)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Intra-workbook reference parser
# ---------------------------------------------------------------------------

def _parse_intra_refs(
    formula: str,
    current_sheet: str,
) -> list[tuple[str, str]]:
    """Parse a formula and return all intra-workbook cell references.

    Returns list of (sheet_name, cell_ref_or_range).
    Excludes any references that overlap with external workbook references.
    """
    # Build exclusion spans from external reference matches
    ext_spans: list[tuple[int, int]] = []
    for m in _REF_RE.finditer(formula):
        ext_spans.append((m.start(), m.end()))

    refs: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()

    for m in _INTRA_REF_RE.finditer(formula):
        # Skip if this match overlaps with an external reference
        mstart, mend = m.start(), m.end()
        if any(es <= mstart < ee or es < mend <= ee for es, ee in ext_spans):
            continue

        quoted_sheet = m.group(1)
        unquoted_sheet = m.group(2)
        cell_ref = m.group(3).replace("$", "")

        if quoted_sheet:
            sheet = quoted_sheet
        elif unquoted_sheet:
            # Check it's not a function name — function names are followed by (
            # and don't have ! before the cell ref. But the regex requires ! for
            # sheet prefix, so unquoted_sheet here means we matched "Name!A1".
            sheet = unquoted_sheet
        else:
            sheet = current_sheet

        # Validate that cell_ref actually looks like a cell (not e.g. "IF1")
        clean = cell_ref.split(":")[0]
        cm = _CELL_RE.match(clean)
        if not cm:
            continue
        # Sanity check: column letters should be <= XFD (max Excel column)
        col_str = cm.group(1).upper()
        if len(col_str) > 3 or (len(col_str) == 3 and col_str > "XFD"):
            continue

        key = (sheet, cell_ref.upper())
        if key in seen:
            continue
        seen.add(key)
        refs.append((sheet, cell_ref))

    return refs


# ---------------------------------------------------------------------------
# External link index resolver
# ---------------------------------------------------------------------------

def _get_link_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map numeric external link indices to filenames and full paths.

    Returns dict: index_str → (filename, full_target_path).
    Reads xl/externalLinks/_rels/externalLink{n}.xml.rels.
    """
    link_map: dict[str, tuple[str, str]] = {}
    names = set(zf.namelist())

    for name in names:
        m = re.match(r"xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels$", name)
        if not m:
            continue
        idx = m.group(1)
        try:
            rels_xml = zf.read(name)
            root = etree.fromstring(rels_xml)
            for rel in root:
                target = rel.get("Target", "")
                rel_type = rel.get("Type", "")
                if target and "externalLink" in rel_type:
                    filename = _extract_filename(target)
                    if filename:
                        link_map[idx] = (filename, unquote(target))
        except Exception:
            pass

    return link_map


# ---------------------------------------------------------------------------
# Streaming formula scanners
# ---------------------------------------------------------------------------

def _stream_external_formulas(
    data: bytes,
    sheet_name: str,
    cell_filter: CellFilter | None = None,
) -> list[tuple[str, str]]:
    """Stream-parse sheet XML and return (cell_ref, formula_text) for formulas
    that contain external workbook references (the '[' character).

    If *cell_filter* is provided, only formulas in matching cells are returned.
    """
    results: list[tuple[str, str]] = []
    in_cell = False
    cell_ref = ""
    has_formula = False
    formula_text = ""
    skip_cell = False

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
                    formula_text = ""
                    skip_cell = False
                    cell_ref = elem.get("r", "")
                    # Fast filter: check cell_filter before processing
                    if cell_filter and cell_ref:
                        m = _CELL_RE.match(cell_ref)
                        if m:
                            col = _col_to_idx(m.group(1))
                            row = int(m.group(2))
                            if not cell_filter.contains(sheet_name, row, col):
                                skip_cell = True
                                in_cell = False
                elif ltag == "f" and in_cell and not skip_cell:
                    has_formula = True
            else:  # "end"
                if ltag == "f" and in_cell and has_formula:
                    formula_text = elem.text or ""
                elif ltag == "c":
                    if (in_cell and has_formula and formula_text
                            and cell_ref and "[" in formula_text):
                        results.append((cell_ref, formula_text))
                    in_cell = False
                    skip_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()

        del context
    except Exception:
        pass

    return results


def _stream_all_formulas(data: bytes) -> dict[str, str]:
    """Stream-parse sheet XML and return {cell_ref: formula_text} for ALL
    formula cells (not just external references).

    Used to build the formula cache for precedent walking.
    """
    formulas: dict[str, str] = {}
    in_cell = False
    cell_ref = ""
    has_formula = False
    formula_text = ""

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
                    formula_text = ""
                    cell_ref = elem.get("r", "")
                elif ltag == "f" and in_cell:
                    has_formula = True
            else:  # "end"
                if ltag == "f" and in_cell and has_formula:
                    formula_text = elem.text or ""
                elif ltag == "c":
                    if in_cell and has_formula and formula_text and cell_ref:
                        formulas[cell_ref] = formula_text
                    in_cell = False
                    elem.clear()
                elif ltag == "row":
                    elem.clear()

        del context
    except Exception:
        pass

    return formulas


def _load_formula_cache(
    zf: zipfile.ZipFile,
    sheet_map: dict[str, str],
    names: set[str],
) -> dict[str, dict[str, str]]:
    """Load all formulas from all sheets into a cache.

    Returns: {sheet_name: {cell_ref: formula_text}}
    """
    cache: dict[str, dict[str, str]] = {}
    for sheet_name, sheet_path in sheet_map.items():
        if sheet_path not in names:
            continue
        try:
            data = zf.read(sheet_path)
            cache[sheet_name] = _stream_all_formulas(data)
        except Exception:
            cache[sheet_name] = {}
    return cache


# ---------------------------------------------------------------------------
# Formula reference parsing
# ---------------------------------------------------------------------------

def _parse_formula_refs(
    formula: str,
    link_map: dict[str, tuple[str, str]],
) -> list[tuple[str, str, str, str]]:
    """Parse a formula and return all external references.

    Returns list of (filename, sheet_name, cell_range, display_path).
    Numeric indices like [1] are resolved via *link_map*.
    """
    refs: list[tuple[str, str, str, str]] = []
    seen: set[tuple[str, str, str]] = set()

    for m in _REF_RE.finditer(formula):
        file_id = m.group(1)
        sheet = m.group(2).strip()
        cell_range = m.group(3).replace("$", "")

        display_path = ""

        if file_id.isdigit():
            # Numeric index → resolve via link map
            resolved = link_map.get(file_id)
            if resolved:
                filename, display_path = resolved
            else:
                filename = f"[ExternalLink{file_id}]"
                display_path = f"xl/externalLinks/externalLink{file_id}.xml"
        else:
            filename = file_id
            # Try to extract full path from the formula
            for pm in _PATH_RE.finditer(formula):
                if pm.group(2) == file_id:
                    display_path = pm.group(1) + filename
                    break
            if not display_path:
                display_path = filename

        key = (filename.lower(), sheet.lower(), cell_range.upper())
        if key in seen:
            continue
        seen.add(key)
        refs.append((filename, sheet, cell_range, display_path))

    return refs


# ---------------------------------------------------------------------------
# Precedent walking (BFS)
# ---------------------------------------------------------------------------

@dataclass
class _PrecedentHit:
    """An external reference found via transitive precedent walking."""
    # The chain from the starting cell to the cell with the external ref.
    # Each tuple: (sheet_name, cell_ref, formula_snippet).
    # Does NOT include the starting cell itself — only intermediaries.
    chain: list[tuple[str, str, str]]
    # The external refs found at the end of the chain.
    external_refs: list[tuple[str, str, str, str]]  # (filename, sheet, range, display_path)


def _walk_precedents(
    start_sheet: str,
    start_cell: str,
    formula_cache: dict[str, dict[str, str]],
    link_map: dict[str, tuple[str, str]],
    max_depth: int = _MAX_PRECEDENT_DEPTH,
    max_cells: int = _MAX_CELLS_VISITED,
) -> list[_PrecedentHit]:
    """BFS walk of formula precedents from a starting cell.

    Looks for cells whose formulas contain external workbook references
    (the '[' character) by walking the intra-workbook dependency graph.

    Returns a list of _PrecedentHit — one per distinct external reference
    found through the precedent chain.
    """
    hits: list[_PrecedentHit] = []
    visited: set[tuple[str, str]] = {(start_sheet, start_cell.upper())}

    # Queue items: (sheet, cell, chain_so_far, depth)
    # chain_so_far: list of (sheet, cell, formula_snippet) for intermediate cells
    queue: deque[tuple[str, str, list[tuple[str, str, str]], int]] = deque()

    # Seed the queue with the direct precedents of the starting cell
    start_formula = formula_cache.get(start_sheet, {}).get(start_cell, "")
    if not start_formula:
        return hits

    intra_refs = _parse_intra_refs(start_formula, start_sheet)
    for ref_sheet, ref_range in intra_refs:
        for cell in _expand_range(ref_range):
            cell_upper = cell.upper()
            key = (ref_sheet, cell_upper)
            if key not in visited:
                visited.add(key)
                queue.append((ref_sheet, cell_upper, [], 1))

    while queue and len(visited) < max_cells:
        sheet, cell, chain, depth = queue.popleft()

        if depth > max_depth:
            continue

        formula = formula_cache.get(sheet, {}).get(cell, "")
        if not formula:
            continue  # hardcoded value or empty — dead end

        if "[" in formula:
            # Found external reference — record the hit
            ext_refs = _parse_formula_refs(formula, link_map)
            if ext_refs:
                full_chain = chain + [(sheet, cell, formula[:200])]
                hits.append(_PrecedentHit(
                    chain=full_chain,
                    external_refs=ext_refs,
                ))
            # Don't walk past this cell — the external file's internals
            # are handled by the next level of trace_formula_levels
            continue

        # No external ref — walk this cell's precedents
        intra_refs = _parse_intra_refs(formula, sheet)
        next_chain = chain + [(sheet, cell, formula[:200])]

        for ref_sheet, ref_range in intra_refs:
            for next_cell in _expand_range(ref_range):
                next_cell_upper = next_cell.upper()
                key = (ref_sheet, next_cell_upper)
                if key not in visited:
                    visited.add(key)
                    queue.append((ref_sheet, next_cell_upper, next_chain, depth + 1))

    return hits


# ---------------------------------------------------------------------------
# File-level scanner
# ---------------------------------------------------------------------------

def scan_external_refs(
    path: Path,
    source_file: str,
    level: int,
    cell_filter: CellFilter | None = None,
    search_dirs: list[Path] | None = None,
) -> list[ExternalReference]:
    """Scan an Excel file for formulas with external workbook references.

    Parameters
    ----------
    path : Path
        Excel file to scan.
    source_file : str
        Display name for the source file (used in ExternalReference.source_file).
    level : int
        Tracing level (1 = model file, 2 = first upstream, etc.).
    cell_filter : CellFilter | None
        If provided, only scan cells within these ranges (for Level 2+).
        When a cell_filter is set, precedent walking is enabled: if a filtered
        cell has a formula without a direct external reference, its in-workbook
        precedents are traced transitively until an external reference is found.
    search_dirs : list[Path] | None
        Directories to search for referenced files.

    Returns list of ExternalReference.
    """
    refs: list[ExternalReference] = []
    search = search_dirs or [path.parent]

    try:
        with zipfile.ZipFile(path) as zf:
            sheet_map = _get_sheet_map(zf)
            names = set(zf.namelist())
            link_map = _get_link_map(zf)

            if cell_filter is None:
                # ── Level 1: stream-scan for direct external refs only ──
                for sheet_name, sheet_path in sheet_map.items():
                    if sheet_path not in names:
                        continue
                    try:
                        data = zf.read(sheet_path)
                        formulas = _stream_external_formulas(data, sheet_name)
                        for cell_ref, formula in formulas:
                            parsed = _parse_formula_refs(formula, link_map)
                            for filename, tgt_sheet, tgt_range, display_path in parsed:
                                resolved, found = _resolve_file(filename, search)
                                refs.append(ExternalReference(
                                    level=level,
                                    source_file=source_file,
                                    source_sheet=sheet_name,
                                    source_cell=cell_ref,
                                    formula=formula[:300],
                                    target_file=filename,
                                    target_sheet=tgt_sheet,
                                    target_range=tgt_range,
                                    target_path=display_path,
                                    file_found=found,
                                    resolved_path=str(resolved),
                                    precedent_chain=None,
                                ))
                    except Exception:
                        pass
            else:
                # ── Level 2+: load formula cache for precedent walking ──
                formula_cache = _load_formula_cache(zf, sheet_map, names)

                for sheet_name in sheet_map:
                    if not cell_filter.has_sheet(sheet_name):
                        continue

                    sheet_formulas = formula_cache.get(sheet_name, {})

                    for cell_ref, formula in sheet_formulas.items():
                        # Check cell_filter
                        m = _CELL_RE.match(cell_ref)
                        if not m:
                            continue
                        col = _col_to_idx(m.group(1))
                        row = int(m.group(2))
                        if not cell_filter.contains(sheet_name, row, col):
                            continue

                        if "[" in formula:
                            # Direct external reference
                            parsed = _parse_formula_refs(formula, link_map)
                            for filename, tgt_sheet, tgt_range, display_path in parsed:
                                resolved, found = _resolve_file(filename, search)
                                refs.append(ExternalReference(
                                    level=level,
                                    source_file=source_file,
                                    source_sheet=sheet_name,
                                    source_cell=cell_ref,
                                    formula=formula[:300],
                                    target_file=filename,
                                    target_sheet=tgt_sheet,
                                    target_range=tgt_range,
                                    target_path=display_path,
                                    file_found=found,
                                    resolved_path=str(resolved),
                                    precedent_chain=None,
                                ))
                        else:
                            # No direct external ref — walk precedents
                            hits = _walk_precedents(
                                sheet_name, cell_ref,
                                formula_cache, link_map,
                            )
                            for hit in hits:
                                for filename, tgt_sheet, tgt_range, display_path in hit.external_refs:
                                    resolved, found = _resolve_file(filename, search)
                                    refs.append(ExternalReference(
                                        level=level,
                                        source_file=source_file,
                                        source_sheet=sheet_name,
                                        source_cell=cell_ref,
                                        formula=formula[:300],
                                        target_file=filename,
                                        target_sheet=tgt_sheet,
                                        target_range=tgt_range,
                                        target_path=display_path,
                                        file_found=found,
                                        resolved_path=str(resolved),
                                        precedent_chain=hit.chain,
                                    ))

    except zipfile.BadZipFile:
        pass
    except Exception:
        pass

    return refs


def _resolve_file(filename: str, search_dirs: list[Path]) -> tuple[Path, bool]:
    """Find a file by name in the search directories.

    Returns (path, found). If not found, returns the first candidate path.
    """
    for d in search_dirs:
        candidate = d / filename
        if candidate.exists():
            return candidate.resolve(), True

    # Case-insensitive fallback
    fn_lower = filename.lower()
    for d in search_dirs:
        if not d.is_dir():
            continue
        try:
            for p in d.iterdir():
                if p.name.lower() == fn_lower:
                    return p.resolve(), True
        except Exception:
            pass

    # Not found — return expected path for display
    return search_dirs[0] / filename if search_dirs else Path(filename), False


# ---------------------------------------------------------------------------
# Recursive level tracer
# ---------------------------------------------------------------------------

def trace_formula_levels(
    model_path: Path,
    search_dirs: list[Path] | None = None,
    max_level: int = 10,
    verbose: bool = False,
) -> dict[int, list[ExternalReference]]:
    """Recursively trace external formula references through multiple levels.

    Level 1: all external references in the model file.
    Level N: external references found in Level N-1's target cells
             (including transitive precedent walking).

    Stops when no new upstream files are found, files don't exist, or
    *max_level* is reached.

    Returns dict[level, list[ExternalReference]].
    """
    if search_dirs is None:
        search_dirs = [model_path.parent]

    results: dict[int, list[ExternalReference]] = {}
    visited_files: set[str] = {model_path.name.lower()}

    # ── Level 1: scan entire model file ────────────────────────────────
    if verbose:
        print(f"\nFormula tracing Level 1: scanning {model_path.name} ...")

    level_1 = scan_external_refs(
        model_path, model_path.name, level=1,
        cell_filter=None,
        search_dirs=search_dirs,
    )
    if verbose:
        n_found = sum(1 for r in level_1 if r.file_found)
        n_missing = sum(1 for r in level_1 if not r.file_found)
        print(f"  Level 1: {len(level_1)} external refs "
              f"({n_found} files found, {n_missing} missing)")

    if not level_1:
        return results

    results[1] = level_1

    # ── Levels 2+ ──────────────────────────────────────────────────────
    for level in range(2, max_level + 1):
        prev_refs = results[level - 1]

        # Group by target file: filename → [(target_sheet, target_range)]
        targets: dict[str, list[tuple[str, str]]] = {}
        target_paths: dict[str, Path] = {}

        for ref in prev_refs:
            if not ref.file_found:
                continue
            fn_lower = ref.target_file.lower()
            if fn_lower in visited_files:
                continue  # avoid cycles
            targets.setdefault(ref.target_file, []).append(
                (ref.target_sheet, ref.target_range),
            )
            target_paths[ref.target_file] = Path(ref.resolved_path)

        if not targets:
            break

        level_refs: list[ExternalReference] = []

        for filename, cell_specs in targets.items():
            file_path = target_paths[filename]
            visited_files.add(filename.lower())

            # Build cell filter from the cell specs referenced at the previous level
            cell_filter = CellFilter.from_refs(cell_specs)

            if verbose:
                n_sheets = len(cell_filter.ranges)
                n_rects = sum(len(r) for r in cell_filter.ranges.values())
                print(f"  Level {level}: scanning {filename} "
                      f"({n_sheets} sheet(s), {n_rects} range(s)) ...")

            file_refs = scan_external_refs(
                file_path, filename, level=level,
                cell_filter=cell_filter,
                search_dirs=search_dirs,
            )
            level_refs.extend(file_refs)

        if not level_refs:
            if verbose:
                print(f"  Level {level}: no further external references found")
            break

        if verbose:
            n_found = sum(1 for r in level_refs if r.file_found)
            n_missing = sum(1 for r in level_refs if not r.file_found)
            n_transitive = sum(1 for r in level_refs if r.precedent_chain)
            print(f"  Level {level}: {len(level_refs)} external refs "
                  f"({n_found} found, {n_missing} missing, "
                  f"{n_transitive} via precedent chain)")

        results[level] = level_refs

    return results
