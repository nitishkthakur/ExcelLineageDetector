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

5. REPORT
   → upstream_tracing_<name>.xlsx
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

The output Excel report (`upstream_tracing_<name>.xlsx`) has two sheets:

### Config Sheet
Lists all configuration parameters, model file, upstream files, and summary statistics.

### Tracing Results Sheet

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
