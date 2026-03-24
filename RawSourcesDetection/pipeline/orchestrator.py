"""Core orchestrator for RawSourcesDetection pipeline.

Reuses existing lineage modules — no re-implementation:
  - lineage.tracing.formula_tracer.scan_external_refs  → formula-level tracing
  - lineage.tracing.tracer.UpstreamTracer               → parallel vector matching
  - lineage.detector.ExcelLineageDetector               → ODBC/OLE DB/PQ harvesting

Performance design:
  - Formula scan: lxml iterparse O(n) streaming, constant memory per file
  - Vector matching: numpy batched hash lookup (exact) or matrix multiply (approx)
  - Upstream file scanning: ProcessPoolExecutor via UpstreamTracer
  - Connection harvest: all files scanned sequentially (I/O bound, reasonable)
"""
from __future__ import annotations

import time
from pathlib import Path

from .config import RSDConfig
from .models import (
    DetectionResult,
    FormulaRef,
    MatchedVector,
    MissingFile,
    RawSource,
    SourceNode,
    UnmatchedVector,
)


# ---------------------------------------------------------------------------
# Step 1: Formula dependency tracing
# ---------------------------------------------------------------------------

def _trace_formulas(
    model_path: Path,
    model_sheets: list[str],
    search_dirs: list[Path],
    max_level: int,
    verbose: bool,
) -> tuple[list[SourceNode], list[FormulaRef], list[MissingFile]]:
    """Recursively scan files for external formula references.

    Level 1: scan model file, keep only refs from user-specified sheets.
    Level 2+: scan each found upstream file, ALL sheets (no restriction).

    Uses lxml iterparse O(n) streaming via scan_external_refs.
    Stops at max_level or when no new files are discovered.
    """
    from lineage.tracing.formula_tracer import scan_external_refs

    source_nodes: list[SourceNode] = [
        SourceNode(model_path.name, str(model_path), level=0, found_on_disk=True)
    ]
    formula_refs: list[FormulaRef] = []
    missing_map: dict[str, MissingFile] = {}   # lower(filename) → MissingFile
    visited: set[str] = {model_path.name.lower()}  # cycle guard

    # Queue: (file_path, scan_level)
    queue: list[tuple[Path, int]] = [(model_path, 1)]

    while queue:
        file_path, scan_level = queue.pop(0)

        if verbose:
            print(f"  [Level {scan_level}] Scanning {file_path.name} ...")

        refs = scan_external_refs(
            path=file_path,
            source_file=file_path.name,
            level=scan_level,
            cell_filter=None,           # full streaming scan — O(n) constant memory
            search_dirs=search_dirs,
        )

        # Level 1 (model file only): filter to user-specified sheets
        if scan_level == 1 and model_sheets:
            sheet_set = set(model_sheets)
            refs = [r for r in refs if r.source_sheet in sheet_set]

        # Dedup within this file's scan: (sheet, cell, target_file, sheet, range)
        seen_keys: set[tuple] = set()
        for ref in refs:
            key = (
                ref.source_sheet, ref.source_cell,
                ref.target_file.lower(), ref.target_sheet.lower(),
                ref.target_range.upper(),
            )
            if key in seen_keys:
                continue
            seen_keys.add(key)

            formula_refs.append(FormulaRef(
                level=scan_level,
                source_file=ref.source_file,
                source_sheet=ref.source_sheet,
                source_cell=ref.source_cell,
                target_file=ref.target_file,
                target_sheet=ref.target_sheet,
                target_range=ref.target_range,
                file_found=ref.file_found,
                resolved_path=ref.resolved_path,
            ))

            target_lower = ref.target_file.lower()

            if ref.file_found:
                resolved = Path(ref.resolved_path)
                if target_lower not in visited:
                    visited.add(target_lower)
                    source_nodes.append(SourceNode(
                        filename=ref.target_file,
                        path=str(resolved),
                        level=scan_level,
                        found_on_disk=True,
                    ))
                    # Recurse if within depth limit
                    if scan_level < max_level:
                        queue.append((resolved, scan_level + 1))
            else:
                # Missing file: accumulate details
                if target_lower not in missing_map:
                    missing_map[target_lower] = MissingFile(
                        filename=ref.target_file,
                        level=scan_level,
                        referenced_by=ref.source_file,
                    )
                mf = missing_map[target_lower]
                if ref.target_sheet and ref.target_sheet not in mf.sheets_needed:
                    mf.sheets_needed.append(ref.target_sheet)
                loc = (
                    f"{ref.source_sheet}!{ref.source_cell}"
                    if ref.source_cell else ref.source_sheet
                )
                if loc and loc not in mf.cells_referencing:
                    mf.cells_referencing.append(loc)

    return source_nodes, formula_refs, list(missing_map.values())


# ---------------------------------------------------------------------------
# Step 2: Hardcoded vector matching
# ---------------------------------------------------------------------------

def _match_vectors(
    model_path: Path,
    model_sheets: list[str],
    input_paths: list[Path],
    config: RSDConfig,
    verbose: bool,
) -> tuple[list[MatchedVector], list[UnmatchedVector]]:
    """Match hardcoded numeric vectors in model sheets against all input files.

    Reuses UpstreamTracer which handles:
      - ProcessPoolExecutor parallel scanning of all upstream files
      - Numpy batched exact matching (hash O(1) + sliding_window for subsequences)
      - Optional numpy matrix-multiply approximate matching (Pearson/cosine/Euclidean)
    """
    if not input_paths:
        return [], []

    from lineage.tracing.tracer import UpstreamTracer

    trace_cfg = config.to_trace_config()
    tracer = UpstreamTracer(trace_cfg, verbose=verbose)

    all_matched: list[MatchedVector] = []
    all_unmatched: list[UnmatchedVector] = []

    for sheet in model_sheets:
        matches, unmatched = tracer.trace(
            model_path=model_path,
            sheet_name=sheet,
            upstream_paths=input_paths,
        )
        for m in matches:
            all_matched.append(MatchedVector(
                model_sheet=m.model_sheet,
                model_range=m.model_range,
                model_length=m.model_length,
                model_sample=m.model_sample,
                match_type=m.match_type,
                similarity=m.similarity,
                upstream_file=m.upstream_file,
                upstream_sheet=m.upstream_sheet,
                upstream_range=m.upstream_range,
                upstream_sample=m.upstream_sample,
            ))
        for u in unmatched:
            all_unmatched.append(UnmatchedVector(
                model_sheet=u.sheet,
                model_range=u.cell_range,
                model_length=u.length,
                model_sample=list(u.values[:5]),
            ))

    return all_matched, all_unmatched


# ---------------------------------------------------------------------------
# Step 3: Connection harvesting (ODBC / OLE DB / Power Query / etc.)
# ---------------------------------------------------------------------------

def _harvest_connections(files: list[Path], verbose: bool) -> list[RawSource]:
    """Run ExcelLineageDetector on every file and collect unique connections.

    Deduplicates by (connection_string, sub_type, source_file) to avoid
    reporting the same ODBC string from 10 formulas in one sheet.
    """
    from lineage.detector import ExcelLineageDetector

    detector = ExcelLineageDetector()
    seen: set[str] = set()
    raw: list[RawSource] = []

    for file_path in files:
        if not file_path.exists():
            continue
        if verbose:
            print(f"  Harvesting {file_path.name} ...")
        try:
            conns = detector.detect(file_path)
        except Exception:
            continue
        for c in conns:
            key = f"{c.raw_connection}|{c.sub_type}|{file_path.name}"
            if key in seen:
                continue
            seen.add(key)
            raw.append(RawSource(
                source_file=file_path.name,
                category=c.category,
                sub_type=c.sub_type,
                connection=c.raw_connection,
                location=c.location,
            ))

    return raw


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def run(
    model_path: Path,
    inputs_dir: Path,
    config: RSDConfig,
    verbose: bool = False,
) -> DetectionResult:
    """Run the full RawSourcesDetection pipeline.

    Steps:
    1. Formula tracing — recursively scan for external file references up to
       max_formula_levels.  Level 1 is scoped to model_sheets; Level 2+ scans
       all sheets of each found upstream file.
    2. Vector matching — scan model sheet(s) for hardcoded numeric vectors and
       match them (exact / approximate) against ALL found input files using
       numpy-vectorized matchers via UpstreamTracer.
    3. Connection harvesting — run ExcelLineageDetector on model + all found
       inputs to extract ODBC, OLE DB, Power Query, and other connections.
    """
    t0 = time.perf_counter()

    # search_dirs: where to look when resolving referenced file names
    search_dirs = [inputs_dir, model_path.parent]

    if verbose:
        print(f"\n=== RawSourcesDetection ===")
        print(f"Model     : {model_path.name}")
        print(f"Sheets    : {config.model_sheets or '(all)'}")
        print(f"Inputs dir: {inputs_dir}")
        print(f"Max levels: {config.max_formula_levels}")
        print(f"Matching  : {'exact+approx' if config.approximate else 'exact only'}")

    # ── Step 1: Formula tracing ──────────────────────────────────────────────
    t1 = time.perf_counter()
    if verbose:
        print(f"\n[Step 1] Formula dependency tracing ...")

    source_nodes, formula_refs, missing_files = _trace_formulas(
        model_path, config.model_sheets, search_dirs,
        config.max_formula_levels, verbose,
    )

    if verbose:
        n_found = sum(1 for n in source_nodes if n.found_on_disk and n.level > 0)
        print(
            f"  {n_found} files found, {len(missing_files)} missing, "
            f"{len(formula_refs)} formula refs  [{time.perf_counter() - t1:.2f}s]"
        )

    # ── Step 2: Vector matching ──────────────────────────────────────────────
    t2 = time.perf_counter()
    if verbose:
        print(f"\n[Step 2] Hardcoded vector matching ...")

    input_paths = [
        Path(n.path) for n in source_nodes if n.found_on_disk and n.level > 0
    ]
    matched_vectors, unmatched_vectors = _match_vectors(
        model_path, config.model_sheets, input_paths, config, verbose,
    )

    if verbose:
        print(
            f"  {len(matched_vectors)} matched, {len(unmatched_vectors)} unmatched"
            f"  [{time.perf_counter() - t2:.2f}s]"
        )

    # ── Step 3: Connection harvesting ────────────────────────────────────────
    t3 = time.perf_counter()
    if verbose:
        print(f"\n[Step 3] Connection harvesting ...")

    all_files = [model_path] + input_paths
    raw_sources = _harvest_connections(all_files, verbose)

    if verbose:
        print(
            f"  {len(raw_sources)} unique connections"
            f"  [{time.perf_counter() - t3:.2f}s]"
        )
        print(f"\nTotal: {time.perf_counter() - t0:.2f}s")

    return DetectionResult(
        model_file=model_path.name,
        source_nodes=source_nodes,
        formula_refs=formula_refs,
        missing_files=missing_files,
        matched_vectors=matched_vectors,
        unmatched_vectors=unmatched_vectors,
        raw_sources=raw_sources,
    )
