# Upstream Tracing Algorithm

This document describes the algorithm used by `trace_upstream.py` to identify the original source of hardcoded vectors in a model Excel file.

---

## Problem Statement

Financial analysts build model spreadsheets by copy-pasting values from upstream data sources (other Excel files, Bloomberg exports, database dumps). Over time, the lineage — where each number came from — is lost. This tool reconstructs that lineage by matching hardcoded value sequences in a **model file** against numeric sequences in a set of **upstream files**.

---

## Key Definitions

| Term | Definition |
|------|-----------|
| **Model file** | The Excel file being analysed — the "destination" of copy-pasted values |
| **Upstream file** | A potential source file from which values may have been copied |
| **Vector** | A contiguous run of numeric cells in a single row or column (minimum length 3) |
| **Hardcoded cell** | A cell containing a numeric value with no formula (`<v>` present, `<f>` absent in the XML) |
| **Model vector** | A hardcoded vector in the model file — candidate for tracing |
| **Upstream vector** | A numeric vector in an upstream file — may contain both hardcoded AND formula-derived values |

---

## Pipeline Overview

```
1. SCAN MODEL SHEET
   Model file + target sheet
     → lxml iterparse (hardcoded numerics only)
     → group into vectors (contiguous runs in same row/column)
     → list[TracingVector]

2. SCAN UPSTREAM FILES  (parallel, ProcessPoolExecutor)
   Each upstream .xlsx/.xlsm file
     → lxml iterparse (ALL numerics, including formula results)
     → group into vectors
     → list[TracingVector]

3. EXACT MATCHING  (hash index + batched numpy sliding window)
   For each model vector:
     a) Hash lookup → O(1) full-vector match
     b) Batched subsequence scan → numpy vectorized

4. APPROXIMATE MATCHING  (batched numpy similarity kernels)
   For each model vector:
     → batch-compare against upstream vectors grouped by length
     → sliding-window for length-mismatched vectors
     → top-N by similarity score

5. RECURSIVE FORMULA TRACING  (streaming XML, scoped by CellFilter)
   Model file (all cells)
     → Level 1: find all formulas referencing external workbooks
     → Level 2+: scan only referenced cell ranges in upstream files
     → stop when file not found on disk or max_level reached

6. REPORT
   → upstream_tracing_<name>.xlsx
     (Config + Tracing Results + Level 1, Level 2, ...)

7. MERMAID VISUALISATION  (optional, trace_upstream_mermaid.py)
   → reads Level N sheets from the report
   → aggregates edges (source sheet → target sheet, labelled with cell ranges)
   → outputs *_mermaid.md with fenced mermaid code block
```

---

## Phase 1: Scanning

### Model File (Hardcoded Only)

The model scanner reads the target sheet's XML using `lxml.etree.iterparse` (streaming, O(n) time, constant memory). For each `<c>` (cell) element:

- If a `<f>` (formula) child is present → skip
- If a `<v>` (value) child is present with numeric type → keep

This captures only values that were typed or pasted, not formula results.

### Upstream Files (All Numerics)

The upstream scanner is identical except it captures ALL numeric `<v>` values regardless of whether a `<f>` element is present. This is critical because:

- An analyst may copy the **result** of a formula from an upstream file and paste it as a value in the model
- The cached value in the XML's `<v>` element contains the last-computed result
- We need to match against these formula results

### Vector Construction

After scanning, cells are grouped into vectors:

1. **Column vectors**: Group cells by column index. Sort by row. Find maximal consecutive-row runs of length >= `min_vector_length`.
2. **Row vectors**: Group cells by row index. Sort by column. Find maximal consecutive-column runs of length >= `min_vector_length`.

A single cell can belong to both a column vector and a row vector. Both are reported.

### Parallel Scanning

Upstream files are scanned in parallel using `concurrent.futures.ProcessPoolExecutor`:
- Each file scan is fully independent (no shared state)
- Worker count: `min(cpu_count, n_files)` by default, configurable via `max_workers`
- Results are serialized as dicts for cross-process pickling

---

## Phase 2: Exact Matching

Exact matching finds upstream vectors (or sub-ranges within longer upstream vectors) that contain the **identical** sequence of values as a model vector.

### Full-Vector Match (Hash Lookup)

1. **Index construction**: For each upstream vector, round all values to `exact_decimal_places` (default 8), convert to tuple, and store in a hash map: `tuple(rounded_values) → list[TracingVector]`.

2. **Lookup**: For each model vector, compute the same rounded tuple and look up in the hash map. O(1) per vector.

### Subsequence Match (Batched Sliding Window)

A model vector of length L may appear as a contiguous subsequence within a longer upstream vector of length M > L. To find these:

1. **Group upstream vectors by length**. Pre-stack each group into a 2D numpy array and pre-round to `exact_decimal_places`.

2. For each model vector of length L, for each upstream length group M > L:
   - Apply `np.lib.stride_tricks.sliding_window_view(batch, L, axis=1)` to get a 3D array of shape `(N, M-L+1, L)` — all sliding windows for all N vectors in the group.
   - Compare against the rounded model vector using `np.all(windows == model_rounded, axis=2)` → boolean mask of shape `(N, M-L+1)`.
   - Extract matches from the mask with `np.where()`.

3. **Memory guard**: If the total number of elements `N × (M-L+1) × L` exceeds 10M, fall back to per-vector processing.

### Floating-Point Tolerance

Values are rounded to `exact_decimal_places` (default 8) before comparison. This handles precision loss from Excel's floating-point representation (e.g., `1234.5678` stored as `1234.56780001`).

---

## Phase 3: Approximate Matching

Approximate matching finds upstream vectors that are **similar but not identical** to model vectors — accounting for rounding, scaling, or partial overlap.

### Similarity Metrics

Three metrics are supported (configurable via `similarity_metric`):

| Metric | Formula | Properties |
|--------|---------|-----------|
| **Pearson correlation** (default) | `r = Σ(xᵢ - x̄)(yᵢ - ȳ) / √(Σ(xᵢ - x̄)² · Σ(yᵢ - ȳ)²)` | Scale and shift invariant. Detects "same shape" regardless of units. Best for financial data where values may be in different currencies or scaled. |
| **Cosine similarity** | `cos(θ) = (x · y) / (‖x‖ · ‖y‖)` | Scale invariant but NOT shift invariant. Good when zero-centering is meaningful. |
| **Euclidean similarity** | `1 / (1 + d_norm)` where `d_norm = ‖x-y‖ / (σ · √L)` | Penalises magnitude differences. Normalised by overall scale and vector length. |

### Batch Computation

All similarity computations are vectorized with numpy:

```python
# Pearson correlation: model (L,) vs batch (N, L) → (N,)
model_c = model - model.mean()
batch_c = batch - batch.mean(axis=1, keepdims=True)
numerator = batch_c @ model_c          # single matrix multiply
denominator = sqrt(sum(batch_c², axis=1) * sum(model_c²))
result = numerator / denominator
```

A single matrix multiplication replaces N individual correlation computations.

### Length-Mismatched Matching

Model and upstream vectors may have different lengths (e.g., the analyst copied only 10 of 20 upstream values). Three cases:

**Case 1: Same length (M == L)**
Direct batch comparison. One kernel call per length group.

**Case 2: Upstream longer (M > L)**
Sliding window: use `sliding_window_view(batch, L, axis=1)` on the upstream batch to generate all L-length windows. Reshape to `(N × (M-L+1), L)` and compute similarity in a single kernel call. Take per-vector maximum.

**Case 3: Upstream shorter (M < L)**
Slide upstream over model: generate `(L-M+1)` windows of length M from the model vector. For each window, batch-compare against all upstream vectors of length M. Take per-vector maximum across windows.

Length tolerance is controlled by `length_tolerance_pct` (default 50%): model length 10 matches upstream lengths 5–15. For sliding-window on longer vectors, the upper bound is doubled (to 30 in this example) since the window handles length mismatch.

### Direction Sensitivity

By default (`direction_sensitive: false`), a column vector in the model can match a row vector in the upstream (supporting transposed paste). When enabled, only same-direction matches are considered.

### Top-N Selection

After computing all candidate matches:
1. Sort by similarity descending
2. Take top `top_n` (default 5)
3. Exclude upstream vectors already matched exactly (to avoid duplicate reporting)

---

## Phase 4: Report

The output Excel report (`upstream_tracing_<name>.xlsx`) contains:

### Config Sheet
Lists all configuration parameters, model file, upstream files, and summary statistics.

### Tracing Results Sheet (value-based matching)

| Column | Description |
|--------|-------------|
| Model Range | Cell range in model sheet (e.g., `B3:B15`) |
| Model Direction | `column` or `row` |
| Model Length | Number of cells |
| Model Sample Values | First 5 values |
| Match Rank | 1-based rank (exact matches first, then approximate by similarity) |
| Match Type | `exact`, `exact_subsequence`, or `approximate` |
| Similarity | 1.0 for exact, <1.0 for approximate |
| Upstream File | Source filename |
| Upstream Sheet | Source sheet name |
| Upstream Range | Full upstream vector range |
| Upstream Matched Range | Specific sub-range that matched (for subsequences) |
| Upstream Direction | `column` or `row` |
| Upstream Length | Number of cells in upstream vector |
| Upstream Sample Values | First 5 values |

Color coding:
- **Green**: exact match
- **Light green**: exact subsequence match
- **Blue** (gradient): approximate match (darker = higher similarity)
- **Yellow**: unmatched model vector

### Level N Sheets (formula-based tracing)

If formula tracing is enabled (default), one sheet per recursion level is appended:

| Column | Description |
|--------|-------------|
| Source File | File containing the formula |
| Source Sheet | Sheet containing the formula |
| Source Cell | Cell reference (e.g., `A1`) |
| Formula | Formula text (truncated to 300 chars) |
| Target File | Referenced workbook filename |
| Target Sheet | Referenced sheet name |
| Target Range | Referenced cell range |
| Target Path | Full path from formula/rels |
| File Found | Whether target file exists on disk |
| Resolved Path | Actual disk path if found |

Color coding:
- **Green**: target file found on disk
- **Red/salmon**: target file NOT found on disk (chain stops here)

### Mermaid Flowchart (`trace_upstream_mermaid.py`)

A standalone post-processing script that reads the Level N sheets from the report and produces a Mermaid flowchart diagram. The diagram shows the upstream connection graph at file/sheet/range granularity — suitable for sharing with non-technical stakeholders.

**Algorithm:**
1. Read all "Level N" sheets from the report using openpyxl (read-only mode)
2. Aggregate edges: group source cells that point to the same (src_file, src_sheet) → (tgt_file, tgt_sheet, tgt_range) into a single edge with a compact label
3. Build subgraph per file (containing sheet nodes), with `classDef missing` styling for files not found on disk
4. Output a fenced mermaid code block to a `.md` file

**Design choices:**
- Files are subgraphs, sheets are nodes — keeps the diagram clean at the right level of abstraction
- Edge labels show `"source_cells → target_range"` — enough detail without overwhelming
- Multiple source cells are aggregated (up to 3 shown, then `"... (N cells)"`)
- Red/green styling matches the Excel report colours
- Supports `--lr` for left-to-right layout (useful for wide chains) and `TB` (default, better for deep chains)

---

## Complexity Analysis

| Phase | Time | Memory |
|-------|------|--------|
| Scan 1 file (~200KB) | ~50ms | ~1MB peak (streaming) |
| Scan 40 files parallel (8 workers) | ~300ms | ~40MB peak |
| Build exact hash index | O(Σ vector_lengths) | O(Σ vector_lengths) |
| Exact hash lookup (per model vector) | O(1) | O(1) |
| Exact subsequence (batched, per length group) | O(N × W × L) numpy | O(N × W × L) |
| Approximate same-length (per group) | O(N × L) numpy | O(N × L) |
| Approximate sliding-window (per group) | O(N × W × L) numpy | O(N × W × L) |
| Total matching (all model vectors) | ~0.5–2s typical | ~50–200MB peak |
| Write report | ~100–500ms | ~10MB |

Where N = vectors in a length group, W = number of sliding windows, L = vector length.

### Benchmark

Tested on Indigo.xlsx "Assumptions Sheet" vs 9 upstream XLSX files:

| Phase | Time |
|-------|------|
| Model scan (91 vectors) | <10ms |
| Upstream scan (3,196 vectors, 9 files, parallel) | 220ms |
| Exact matching (hash + subsequence) | ~80ms |
| Approximate matching (Pearson, batched) | ~400ms |
| Report generation | ~200ms |
| **Total** | **~920ms** |

---

## Configuration Reference

All settings are in `tracing_config.json`:

```json
{
  "matching": {
    "exact": true,                    // enable exact matching
    "approximate": true,              // enable approximate matching
    "top_n": 5,                       // top-N approximate matches per vector
    "exact_decimal_places": 8,        // rounding for exact comparison
    "subsequence_matching": true,     // match model as subsequence of longer upstream
    "similarity_metric": "pearson",   // pearson | cosine | euclidean
    "min_similarity": 0.8,            // minimum similarity to report
    "length_tolerance_pct": 50.0,     // allow ±50% length mismatch (approximate)
    "direction_sensitive": false       // allow column↔row matching
  },
  "performance": {
    "max_workers": null,              // null = auto (cpu_count)
    "min_vector_length": 3            // minimum vector length
  }
}
```

### Config Interactions

- If `exact=true, approximate=false`: only exact matches reported, `top_n` is ignored
- If `exact=false, approximate=true`: only approximate matches, no hash index built
- If `exact=true, approximate=true` (default): exact matches shown first, then top-N approximate (excluding already-exact-matched upstream vectors)
- `subsequence_matching` only applies to exact matching; approximate matching always supports length mismatch via sliding window

---

## Phase 5: Recursive Formula Tracing

In addition to value-based matching, the tracer follows **formulas that reference external workbooks** through multiple levels. This is separate from value matching and does not involve similarity comparison — formulas directly point to specific cells in specific files.

### Formula Reference Patterns

External workbook references in Excel formulas take two forms:

| Pattern | Example | Resolution |
|---------|---------|-----------|
| Literal filename | `'[budget.xlsx]Revenue'!A1:A10` | Filename extracted directly |
| Numeric index | `[1]Sheet1!B5` | Index `[n]` → `xl/externalLinks/externalLink{n}.xml` → `.rels` file → actual path |

The regex `_REF_RE` captures both forms, handling optional path prefixes (local, UNC, SharePoint URLs) before the `[filename]` bracket.

### Recursive Scoping

**Level 1** scans the entire model file — all sheets, all cells — for formulas containing external workbook references. Uses streaming XML (constant memory).

**Level 2+** uses `CellFilter` to scope scanning to only the cell ranges identified at the previous level:

```
Level 1: model.xlsx (all cells)
  → finds [upstream.xlsx]Data!A1:A10 in cell C5

Level 2: upstream.xlsx (only Data!A1:A10)
  → A1 = B1 * 2 (no direct external ref)
  → walks precedents: B1 = C1 + 100, C1 = '[source.xlsx]Raw'!D5
  → reports source.xlsx with precedent chain A1 → B1 → C1

Level 3: source.xlsx (only Raw!D5)
  → file not found → stop, flag as missing
```

### Transitive Precedent Walking

At Level 2+, a target cell may not directly reference an external file — its formula may depend on other cells within the same workbook that do. The tracer handles this via **BFS precedent walking**:

1. For each filtered cell that has a formula **without** `[` (no direct external ref):
   a. Parse intra-workbook cell references from the formula using `_INTRA_REF_RE`
   b. Add those cells to a BFS queue
   c. For each queued cell:
      - If its formula contains `[`: parse external refs, record hit with full precedent chain
      - Else: parse its intra-refs and continue walking
   d. Track visited cells to prevent infinite loops from circular references

**Key features:**
- **Cross-sheet support**: `Sheet2!A1` references are followed across sheets within the same workbook
- **Range expansion**: `SUM(B1:B10)` expands to individual cells B1, B2, ..., B10 (capped at 10,000 cells)
- **Arbitrary depth**: Follows chains up to 20 hops deep (configurable via `_MAX_PRECEDENT_DEPTH`)
- **Circular reference safety**: `visited` set prevents infinite BFS loops
- **Multiple paths**: If a cell depends on multiple precedents that each lead to different external files, all are reported
- **Memory guard**: Max 10,000 cells visited per starting cell (`_MAX_CELLS_VISITED`)

**Implementation details:**
- Level 1 uses `_stream_external_formulas` (streaming, constant memory) — only cells with `[` are captured
- Level 2+ uses `_load_formula_cache` to load ALL formulas from ALL sheets into `dict[str, dict[str, str]]` (sheet → cell → formula). This enables O(1) random access for precedent lookups. Memory: ~5MB for 100K formula cells.
- Intra-workbook references are parsed by `_INTRA_REF_RE`, which captures `A1`, `$B$5`, `Sheet2!A1:B10`, `'Sheet Name'!C3` patterns. Matches overlapping with `_REF_RE` (external reference) spans are excluded.

### Cycle Prevention

A `visited_files` set tracks all files already scanned across levels. If a file appears at multiple levels (e.g., A → B → A), it is only scanned once. This prevents infinite recursion in circular reference chains between files.

Within a single file, BFS precedent walking uses a per-cell `visited` set to prevent infinite loops from circular formula references (e.g., A1=B1, B1=A1).

### Data Model

Each external reference is stored as an `ExternalReference`:

| Field | Description |
|-------|-------------|
| `level` | Tracing level (1, 2, 3, ...) |
| `source_file` | File containing the formula |
| `source_sheet` | Sheet containing the formula |
| `source_cell` | Cell reference (e.g., `A1`) — the cell in the CellFilter range |
| `formula` | Formula text of the source cell (truncated to 300 chars) |
| `target_file` | Referenced workbook filename |
| `target_sheet` | Referenced sheet name |
| `target_range` | Referenced cell range (e.g., `A1:A10`) |
| `target_path` | Full path from formula/rels |
| `file_found` | Whether target file exists on disk |
| `resolved_path` | Actual disk path if found |
| `precedent_chain` | `None` for direct references; `list[(sheet, cell, formula)]` for transitive references — the intermediate cells from source_cell to the cell with the external ref |

### File Resolution

1. Search the model file's directory for an exact filename match
2. Fall back to case-insensitive search
3. If not found, return expected path with `file_found=False`

The tracer extracts just the filename from formula paths (which may include local paths, UNC paths, or SharePoint URLs) and searches in the model file's directory.

### Known Limitations

- **Named ranges**: Formulas referencing defined names (e.g., `=Revenue_Total`) are not resolved to cell addresses. This would require parsing `xl/workbook.xml` definedNames.
- **R1C1-style references**: Only A1-style cell references are supported (Excel normalises to A1-style in stored XML, so this rarely matters).
- **Structured table references**: `Table1[Column1]` syntax is not parsed for precedent walking (though it won't be confused with external references).
