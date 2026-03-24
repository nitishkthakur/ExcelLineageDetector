# Business Contract — Algorithm Reference

## Table of Contents

1. [Pipeline Overview](#1-pipeline-overview)
2. [Data Models](#2-data-models)
3. [Model Scanning](#3-model-scanning)
   - 3.1 [Formula Extraction (Streaming XML)](#31-formula-extraction-streaming-xml)
   - 3.2 [Hardcoded Vector Detection](#32-hardcoded-vector-detection)
   - 3.3 [Formula Variable Grouping](#33-formula-variable-grouping)
   - 3.4 [Named Range Scanning](#34-named-range-scanning)
   - 3.5 [Scalar Detection](#35-scalar-detection)
   - 3.6 [Variable Deduplication and Merging](#36-variable-deduplication-and-merging)
4. [Dependency Graph Construction](#4-dependency-graph-construction)
   - 4.1 [Transformation Building (Indexed Overlap)](#41-transformation-building-indexed-overlap)
   - 4.2 [Range Overlap (AABB Intersection)](#42-range-overlap-aabb-intersection)
   - 4.3 [Edge Deduplication](#43-edge-deduplication)
5. [Upstream Lineage Enrichment](#5-upstream-lineage-enrichment)
6. [Formula Conversion (Excel to SQL)](#6-formula-conversion-excel-to-sql)
   - 6.1 [Reference Substitution](#61-reference-substitution)
   - 6.2 [Function Mapping](#62-function-mapping)
   - 6.3 [Nested IF Conversion (Balanced-Paren Parser)](#63-nested-if-conversion-balanced-paren-parser)
7. [LLM Business Name Inference](#7-llm-business-name-inference)
8. [Streaming Cell Reader](#8-streaming-cell-reader)
   - 8.1 [Shared String Resolution](#81-shared-string-resolution)
   - 8.2 [Sheet Summary (Bounded Memory)](#82-sheet-summary-bounded-memory)
   - 8.3 [Cell Neighborhood Extraction](#83-cell-neighborhood-extraction)
9. [Contract Excel Writer](#9-contract-excel-writer)
10. [Mermaid Diagram Generation](#10-mermaid-diagram-generation)
11. [Python Refactor Generation](#11-python-refactor-generation)
    - 11.1 [Topological Sort with Cycle Detection](#111-topological-sort-with-cycle-detection)
12. [Complexity Summary](#12-complexity-summary)

---

## 1. Pipeline Overview

The pipeline converts an Excel model file into a Business Contract — a structured
representation of all quantities, transformations, inputs, outputs, and upstream
data sources in the model.

```
ContractConfig (settings)
    │
    ▼  run.py: generate_contract()
    │
    ├─[1]─► scanner.py: scan_model()
    │       ├─ _scan_connections()          ← ExcelLineageDetector (13 extractors)
    │       ├─ _scan_formulas()             ← streaming XML iterparse
    │       ├─ _scan_hardcoded_vectors()    ← existing hardcoded_scanner
    │       ├─ _scan_formula_variables()    ← grouping + contiguity check
    │       ├─ _scan_named_ranges()         ← definedName elements
    │       ├─ deduplicate + merge names
    │       ├─ _build_transformations()     ← indexed overlap matching
    │       └─ _build_edges()              ← set-based deduplication
    │
    ├─[2]─► graph_builder.py: enrich_upstream()
    │       └─ UpstreamTracer.trace()       ← value-based matching (exact + approx)
    │
    ├─[3]─► formula_converter.py: excel_to_sql()
    │       ├─ reference substitution (regex + dict lookup)
    │       ├─ function mapping (40+ Excel→SQL)
    │       └─ balanced-paren IF→CASE WHEN conversion
    │
    ├─[4]─► llm_namer.py: infer_business_names()
    │       ├─ gather cell neighborhood context
    │       ├─ batch LLM calls (20 vars/batch)
    │       └─ retry + fallback naming
    │
    ├─[5]─► contract_writer.py: write_contract()
    │       └─ 6-sheet Excel (Summary, Variables, Transformations,
    │          Dependencies, External Connections, Upstream Lineage)
    │
    ├─[6]─► mermaid/generator.py: write_mermaid()
    │       ├─ source-level diagram (files → sheets)
    │       └─ variable-level diagram (vars with formula edges)
    │
    └─[7]─► refactor/generator.py: generate_python()
            ├─ topological sort with cycle detection
            └─ code generation (inputs, functions, compute_all)
```

Each step has independent error handling — a failure in any step logs a warning
and continues with partial results.

---

## 2. Data Models

### ContractVariable

Represents a named business quantity — a vector or scalar.

```
Fields:
  id              sha256(sheet | cell_range)[:12]    deterministic, dedup-safe
  business_name   LLM-inferred or fallback           "quarterly_revenue"
  excel_location  "{sheet}!{cell_range}"              "Inputs!B2:B13"
  sheet           sheet name                          "Inputs"
  cell_range      cell range                          "B2:B13"
  direction       "row" | "column" | "scalar"
  length          number of cells in range
  sample_values   first 5 values                     [100.5, 98.3, ...]
  variable_type   "input" | "output" | "intermediate"
  source_type     "hardcoded" | "formula" | "connection" | "external_link"
  upstream_*      populated by graph_builder if matched
  confidence      0.0-1.0 from upstream matching
  match_type      "exact" | "approximate" | None
```

### TransformationStep

One formula transformation in a dependency chain.

```
Fields:
  id                    sha256("tx" | output_id | formula[:50])[:12]
  output_variable_id    which variable this formula computes
  input_variable_ids    list of variable IDs referenced by the formula
  excel_formula         raw Excel formula
  sql_formula           SQL-like converted formula (with business names)
  sheet, cell_range     location of the formula
```

### DependencyEdge

Directed edge in the dependency graph.

```
  source_id     variable or upstream node
  target_id     variable
  edge_type     "formula" | "vector_match" | "external_link" | "connection"
  metadata      dict with edge-specific details
```

### ID Generation

All IDs use truncated SHA-256 for determinism:

```
_make_id(*parts) = sha256("|".join(parts)).hexdigest()[:12]
```

Same inputs always produce the same ID. Collision probability for 12 hex chars
(48 bits) is negligible for workbook-scale data (~1000 variables).

---

## 3. Model Scanning

### 3.1 Formula Extraction (Streaming XML)

**Module:** `scanner.py :: _scan_formulas()`

An `.xlsx` file is a ZIP archive. Each sheet lives at `xl/worksheets/sheet{N}.xml`.
We use `lxml.etree.iterparse` to stream cell elements without loading the full XML
into memory.

```
Algorithm:
  1. Open ZIP, read xl/workbook.xml
  2. Parse <sheet> elements → map sheet_name → rId
  3. Read xl/_rels/workbook.xml.rels → map rId → target path
  4. For each target sheet XML:
     a. iterparse with tag="{ns}c" (cell elements)
     b. For each cell:
        - Extract ref (e.g. "B3"), formula text (<f>), cached value (<v>)
        - If formula exists, emit {sheet, cell, formula, value}
     c. Clear element to release memory

Complexity: O(n) time, O(1) working memory per sheet
            n = total cells across all sheets
```

### 3.2 Hardcoded Vector Detection

**Module:** `scanner.py :: _scan_hardcoded_vectors()`

Delegates to the existing `lineage.hardcoded_scanner.scan_vectors()` which uses
`lxml.etree.iterparse` for O(n) streaming. Detects contiguous runs of non-formula
numeric cells in a single row or column.

```
Filter: vec.length >= min_vector_length (default 3)
Output: ContractVariable with source_type="hardcoded", variable_type="input"
```

### 3.3 Formula Variable Grouping

**Module:** `scanner.py :: _scan_formula_variables()`

Groups individual formula cells into contiguous vector variables.

```
Algorithm:
  1. Parse each formula cell ref → (col_letter, row_num)
  2. Group by (sheet, col_letter) → column candidates
     Group by (sheet, row_num)   → row candidates
  3. For each group with len >= min_length:
     a. Sort by position
     b. Check contiguity: last_pos - first_pos + 1 == count
     c. If contiguous, emit vector variable
     d. Track which cells are consumed (cells_in_vectors set)
  4. Remaining cells → scalar variables (see §3.5)

Contiguity check (column example):
  cells = [B2, B3, B4, B5]  →  rows = [2, 3, 4, 5]
  5 - 2 + 1 == 4 == len(cells)  →  contiguous ✓

  cells = [B2, B3, B5]  →  rows = [2, 3, 5]
  5 - 2 + 1 == 4 ≠ 3  →  not contiguous ✗

Complexity: O(f) where f = formula count
Space: O(f) for grouping dicts
```

### 3.4 Named Range Scanning

**Module:** `scanner.py :: _scan_named_ranges()`

Extracts variables from Excel defined names (`<definedName>` elements in
`xl/workbook.xml`). Financial models heavily use named ranges for assumptions
like `DiscountRate`, `TaxRate`, etc.

```
Algorithm:
  1. Parse workbook.xml to get sheet names in order (for localSheetId mapping)
  2. Parse definedName elements:
     - Skip hidden names and built-in names (_xlnm.*)
     - Skip dynamic names (contain OFFSET, INDIRECT, or parentheses)
     - Regex match: 'SheetName'!$A$1:$B$10 → sheet, start, end
  3. Classify direction:
     - length == 1 → "scalar"
     - more rows than columns → "column"
     - else → "row"
  4. Load sample values via openpyxl (data_only=True)
  5. Set business_name = defined name (already a business name!)

Regex: ^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$

Complexity: O(d + s) where d = defined names, s = sample value reads
```

### 3.5 Scalar Detection

**Module:** `scanner.py :: _scan_formula_variables()` (final loop)

Formula cells that don't belong to any vector are emitted as scalar variables:

```
Algorithm:
  After vector extraction in §3.3:
  for each formula cell (sheet, cell_ref) NOT in cells_in_vectors:
    emit ContractVariable(direction="scalar", length=1, values=[value])

This captures:
  - Single-cell formulas (e.g., =SUM(A1:A10))
  - Isolated calculations (e.g., =B2*C2)
  - Named-range-like scalars defined by formula
```

### 3.6 Variable Deduplication and Merging

**Module:** `scanner.py :: scan_model()`

Named ranges may overlap with hardcoded/formula variables. We deduplicate by
`excel_location` and merge business names from named ranges onto existing variables.

```
Algorithm:
  all_vars = hardcoded_vars + formula_vars
  existing_locations = {v.excel_location for v in all_vars}

  for named_var in named_vars:
    if named_var.excel_location NOT in existing_locations:
      # New variable from named range — add it
      all_vars.append(named_var)
    else:
      # Already exists — merge the business_name onto it
      for v in all_vars:
        if v.excel_location == named_var.excel_location and not v.business_name:
          v.business_name = named_var.business_name

Complexity: O(v + n) where v = existing vars, n = named vars
```

---

## 4. Dependency Graph Construction

### 4.1 Transformation Building (Indexed Overlap)

**Module:** `scanner.py :: _build_transformations()`

Links formula variables to the input variables they reference.

```
Algorithm:
  PRE-INDEX:
    cell_to_var: dict[(sheet, start_cell)] → var_id    O(1) lookup
    sheet_ranges: dict[sheet] → list[(parsed_range, var_id)]

  For each formula f:
    1. Look up output variable via cell_to_var[(f.sheet, f.cell)]
    2. Skip if no variable or already processed (seen_outputs set)
    3. Extract cell/range references from formula using regex:
       (?:'?([^'!(),\[\]]+)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)
    4. For each reference:
       a. Parse to rectangle (min_col, min_row, max_col, max_row)
       b. Check overlap against sheet_ranges[ref_sheet]
       c. If overlapping, add var_id to input_ids (deduplicated via set)
    5. Emit TransformationStep(output_id, input_ids, formula, ...)

  Regex captures:
    Group 1: optional sheet name (excluding parens, commas, brackets)
    Group 2: cell ref with optional range, handles $-signs

  Complexity: O(f × r_avg) where f = formulas, r_avg = ranges per sheet
              Typically r_avg << total_variables since it's per-sheet
  Space: O(v) for indexes
```

### 4.2 Range Overlap (AABB Intersection)

**Module:** `scanner.py :: _ranges_overlap()`

Uses Axis-Aligned Bounding Box intersection to test if two cell ranges overlap.

```
Algorithm:
  1. Parse both ranges to rectangles:
     _parse_range("B2:M13") → (min_col=2, min_row=2, max_col=13, max_row=13)
     _parse_range("B5")     → (min_col=2, min_row=5, max_col=2, max_row=5)

  2. AABB non-overlap test:
     NO overlap if:
       r1.max_col < r2.min_col  OR   (r1 entirely left of r2)
       r1.min_col > r2.max_col  OR   (r1 entirely right of r2)
       r1.max_row < r2.min_row  OR   (r1 entirely above r2)
       r1.min_row > r2.max_row       (r1 entirely below r2)

  3. Overlap = NOT(non-overlap)

  Column index: base-26 encoding
    A→1, B→2, ..., Z→26, AA→27, AB→28, ...

  $-signs stripped before parsing: $B$3 → B3

  Complexity: O(1) time, O(1) space
```

### 4.3 Edge Deduplication

**Module:** `scanner.py :: _build_edges()`

```
Algorithm:
  seen = set()
  for each transformation:
    for each input_id:
      key = (input_id, output_variable_id)
      if key not in seen:
        seen.add(key)
        emit DependencyEdge(source=input_id, target=output_id, type="formula")

  Complexity: O(e) time, O(e) space where e = total input references
```

---

## 5. Upstream Lineage Enrichment

**Module:** `graph_builder.py :: enrich_upstream()`

Traces hardcoded input vectors back to upstream Excel files using value-based
matching (exact hash matching + approximate similarity).

```
Algorithm:
  1. Collect upstream file paths:
     - Explicit paths from config.upstream_files
     - Glob *.xlsx, *.xlsm from config.upstream_dir (recursive)
     - Exclude the model file itself

  2. Configure TraceConfig with similarity settings:
     - similarity_metric (pearson/cosine/euclidean)
     - min_similarity threshold (default 0.8)
     - subsequence_matching (allow partial matches)

  3. For each sheet with hardcoded input variables:
     a. Call UpstreamTracer.trace(model_path, sheet, upstream_paths)
        - Scans model sheet for hardcoded vectors
        - Scans upstream files for ALL numerics (hardcoded + formula)
        - Exact matching: hash-based O(1) lookup + batched numpy subsequence
        - Approximate matching: vectorized Pearson correlation via matrix multiply
     b. Returns (matches, unmatched)

  4. For each match, find the overlapping input variable:
     - Update: upstream_source, upstream_file, upstream_sheet,
               upstream_range, confidence, match_type
     - Change source_type from "hardcoded" to "external_link"
     - Add DependencyEdge(type="vector_match") with match metadata

  Complexity: Dominated by UpstreamTracer
    - Scanning: O(u × file_size) with ProcessPoolExecutor parallelism
    - Exact matching: O(m × u_avg) hash lookups + subsequence scans
    - Approximate matching: O(m × u_avg × L) matrix multiply per length bucket
    where m = model vectors, u = upstream vectors, L = vector length
```

---

## 6. Formula Conversion (Excel to SQL)

**Module:** `formula_converter.py :: excel_to_sql()`

Three-pass conversion: reference substitution → function mapping → IF conversion.

### 6.1 Reference Substitution

```
Algorithm:
  1. Regex: (?:'?\[?([^\]'!(),]+)\]?'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)
  2. For each match:
     - Strip $-signs from reference
     - Construct full_ref = "Sheet!Ref" or just "Ref"
     - Lookup in var_names dict → replace with business name
     - Fallback: keep cleaned reference

  Example:
    Formula:    SUM(Inputs!$B$2:$B$13)
    var_names:  {"Inputs!B2:B13": "revenue"}
    Result:     SUM(revenue)

  Complexity: O(n) where n = formula length
```

### 6.2 Function Mapping

```
Mapping table (40+ entries):
  SUM → SUM          AVERAGE → AVG         VLOOKUP → JOIN_LOOKUP
  IF → CASE WHEN     IFERROR → COALESCE    SUMPRODUCT → SUM_PRODUCT
  STDEV → STDDEV     MID → SUBSTR          TODAY → CURRENT_DATE
  LEN → LENGTH       CONCATENATE → CONCAT  DATEDIF → DATEDIFF
  ...

Algorithm:
  Regex: ([A-Z][A-Z0-9_.]+)\s*\(
  Replace each function name with SQL equivalent from mapping table.
  Unknown functions are kept as-is.
```

### 6.3 Nested IF Conversion (Balanced-Paren Parser)

**Problem:** Simple regex `IF([^,]+,[^,]+,[^)]+)` fails when arguments contain
commas (e.g., `IF(SUMIF(A:A,">0",B:B) > 100, C1, D1)`).

```
Algorithm:
  1. Scan for "CASE WHEN(" markers (IF already renamed by function mapping)
  2. For each marker:
     a. Find matching close paren using depth counter:
        depth = 1
        for each char after opening paren:
          '(' → depth++
          ')' → depth--
          stop when depth == 0
     b. Extract inner content between parens
     c. Split on top-level commas using _split_top_level():

        _split_top_level("SUMIF(A:A,\">0\",B:B) > 100, C1, D1", ",")
                                 ↑ depth=1    ↑ depth=1
        → ["SUMIF(A:A,\">0\",B:B) > 100", " C1", " D1"]
                                            ↑ depth=0 — split here

     d. Emit: CASE WHEN {cond} THEN {then} ELSE {else} END

  _split_top_level(s, sep):
    Track paren depth. Only split on sep when depth == 0.

  Example:
    Input:  IF(SUMIF(A:A,">0",B:B)>100,C1,D1)
    Step 1: CASE WHEN(SUM_IF(A:A,">0",B:B)>100,C1,D1)  [function map]
    Step 2: inner = "SUM_IF(A:A,\">0\",B:B)>100,C1,D1"
    Step 3: parts = ["SUM_IF(A:A,\">0\",B:B)>100", "C1", "D1"]
    Step 4: CASE WHEN SUM_IF(A:A,">0",B:B)>100 THEN C1 ELSE D1 END

  Complexity: O(n) single pass, Space O(n)
```

---

## 7. LLM Business Name Inference

**Module:** `llm_namer.py :: infer_business_names()`

Uses Claude Haiku via LangChain to assign descriptive business names to variables.

```
Algorithm:
  1. Open model file as ZIP, load shared strings once
  2. Process variables in batches of batch_size (default 20):
     a. For each variable in batch:
        - Extract start cell from range
        - Call read_cell_neighborhood(sheet, start_cell, radius=3)
        - Format context: sheet, range, direction, sample_values, nearby cells
     b. Combine all contexts into single prompt
     c. Send to LLM with system prompt:
        - Role: financial analyst
        - Rules: snake_case, self-documenting, include frequency, reflect percentages
        - Output format: JSON array [{id, business_name}]
     d. Parse response:
        - Handle markdown fences (```json ... ```)
        - Validate JSON structure (must be list of dicts)
        - Map id → business_name
     e. Retry logic: up to 2 retries on failure
     f. Fallback: descriptive schema-based names if all retries fail

  3. Final pass: fill any unnamed variables with fallback names

  Fallback naming:
    "{sheet}_{range}_{direction_suffix}"
    e.g., "inputs_b2_to_b13_col_vec", "assumptions_b2_scalar"

  Complexity:
    Context gathering: O(v × cells_per_sheet) per batch
    LLM calls: O(ceil(v / batch_size)) API calls
    Total latency: dominated by LLM response time (~1-5s per batch)
```

---

## 8. Streaming Cell Reader

**Module:** `mcp_server/streaming.py`

All cell reading uses `lxml.etree.iterparse` for constant-memory streaming.

### 8.1 Shared String Resolution

Excel stores repeated strings in a shared string table (`xl/sharedStrings.xml`).
Cells reference strings by index.

```
Algorithm:
  1. iterparse xl/sharedStrings.xml with tag="{ns}si"
  2. For each <si>: concatenate all <t> text children
  3. Append to list (index = position in list)

  Cell resolution:
    If cell type == "s" and value == "42":
      resolved_value = shared_strings[42]

  Complexity: O(s) time and space where s = unique string count
```

### 8.2 Sheet Summary (Bounded Memory)

**Problem:** A sheet with 50,000 rows should not be loaded into memory just
to get headers + 5 sample rows.

```
Algorithm:
  max_keep = 1 + max_rows  (e.g., 6 rows for headers + 5 samples)

  Stream ALL cells:
    - Track max_row, max_col (counters only)
    - Store cell only if cell.row <= max_keep

  After streaming:
    - Extract headers from row 1
    - Extract sample_rows from rows 2..max_keep

  Memory: O(max_keep × max_col) regardless of sheet size
  Time: O(n) — must scan to find max_row/max_col
```

### 8.3 Cell Neighborhood Extraction

```
Algorithm:
  1. Parse center_ref → (crow, ccol)
  2. Define bounding box: [crow-r, ccol-r] to [crow+r, ccol+r]
  3. Stream all cells, collect those within box

  For radius=3: reads a 7×7 grid (up to 49 cells)
  Memory: O(r²) where r = radius
  Time: O(n) — must stream to end of sheet
```

---

## 9. Contract Excel Writer

**Module:** `contract_writer.py :: write_contract()`

Generates a 6-sheet Excel workbook:

```
Sheet 1: Summary
  - Model file path, output sheets
  - Variable counts (total, input, output, intermediate)
  - Transformation and connection counts

Sheet 2: Variables
  - Columns: ID, Business Name, Sheet, Cell Range, Direction, Length,
             Type, Source Type, Sample Values, Upstream Source,
             Confidence, Match Type
  - Color coding: input→green, output→orange, intermediate→blue

Sheet 3: Transformations
  - Columns: ID, Output Variable, Input Variables (comma-separated),
             Excel Formula, SQL Formula, Sheet, Cell Range
  - Variable IDs resolved to business names via var_map

Sheet 4: Dependencies
  - Columns: Source, Target, Edge Type, Details
  - All edges (formula + vector_match + external_link)

Sheet 5: External Connections (conditional — only if connections exist)
  - Columns: ID, Category, Sub-Type, Raw Connection, Location, Confidence
  - From ExcelLineageDetector's 13 extractors

Sheet 6: Upstream Lineage (conditional — only if upstream matches exist)
  - Columns: Business Name, Sheet, Cell Range, Upstream File,
             Upstream Sheet, Upstream Range, Confidence, Match Type

Auto-width: Each column auto-sized to max(cell content length, 10), capped at 50.
```

---

## 10. Mermaid Diagram Generation

**Module:** `mermaid/generator.py`

### Source-Level Diagram

Shows data flow at the file/sheet level.

```
Algorithm:
  1. Collect node sets:
     - upstream_files: from variables with upstream_file + connections
     - model_sheets: all sheets with variables, minus output_sheets
     - output_sheets: from config

  2. Build index maps for O(1) lookup:
     uf_idx = {name: i for i, name in enumerate(sorted(upstream_files))}
     ms_idx = {name: i for i, name in enumerate(sorted(model_sheets))}
     os_idx = {name: i for i, name in enumerate(sorted(output_sheets))}

  3. Generate nodes:
     uf{i}[("filename")]     — upstream (stadium shape)
     ms{i}["sheetname"]      — model (rectangle)
     os{i}[["sheetname"]]    — output (subroutine shape)

  4. Generate edges (deduplicated via seen_edges set):
     - For each variable with upstream: uf → ms or uf → os
     - For each formula edge across sheets: ms → os or ms → ms

  5. Apply CSS styling:
     upstream → green, model → blue, output → pink

  Complexity: O(v + e), no list.index() calls (O(1) dict lookups)

  Example output:
    graph LR
        uf0[("upstream_a.xlsx")]
        ms0["Inputs"]
        ms1["Calculations"]
        os0[["Summary"]]
        uf0 --> ms0
        ms0 --> ms1
        ms1 --> os0
```

### Variable-Level Diagram

Shows individual variables with formula dependency edges.

```
Algorithm:
  1. Partition variables: inputs, intermediates, outputs
  2. Create Mermaid subgraphs for each partition
  3. Node ID = v_{var.id} (guaranteed unique by SHA-256)
  4. Label = sanitized business_name (truncated to 60 chars)
  5. Edges from contract.edges, deduplicated by (source_id, target_id)

  Example output:
    graph TD
        subgraph Inputs
            v_abc123["quarterly_revenue"]
        end
        subgraph Calculations
            v_def456["gross_profit"]
        end
        subgraph Outputs
            v_ghi789[["net_income"]]
        end
        v_abc123 -->|formula| v_def456
        v_def456 -->|formula| v_ghi789
```

---

## 11. Python Refactor Generation

**Module:** `refactor/generator.py :: generate_python()`

Generates a standalone Python module that reproduces the Excel calculation engine.

```
Generated structure:
  1. Input declarations:
     - Vectors → np.zeros(length) with TODO comment
     - Scalars → literal value

  2. Transformation functions (in topological order):
     def compute_{output_name}({input_names}):
         """Excel: {formula}
         SQL:   {sql_formula}"""
         return {placeholder}

  3. Orchestrator:
     def compute_all():
         {output1} = compute_{output1}({inputs})
         {output2} = compute_{output2}({output1}, {inputs})
         ...
         return {"output_name": output_value, ...}

  4. Entry point:
     if __name__ == "__main__":
         results = compute_all()
         for name, value in results.items():
             print(f"{name}: {value}")
```

### 11.1 Topological Sort with Cycle Detection

```
Algorithm: Three-state DFS
  States: UNVISITED(0), IN_PROGRESS(1), DONE(2)

  visit(tx):
    if state[tx] == DONE: return          # already processed
    if state[tx] == IN_PROGRESS: return   # cycle detected — break it
    state[tx] = IN_PROGRESS
    for each input_id in tx.input_variable_ids:
      if input_id has a transformation:
        visit(that transformation)
    state[tx] = DONE
    result.append(tx)                     # post-order = reverse topo

  for each transformation:
    if state == UNVISITED:
      visit(tx)

  Cycle handling:
    When IN_PROGRESS is encountered, the back-edge is dropped.
    The generated Python code will compute the cycle-entry variable
    before its circular dependency is ready — using a placeholder value.

  Complexity: O(t + e) where t = transformations, e = dependency edges
  Space: O(t) for recursion stack + state dict
```

---

## 12. Complexity Summary

| Step | Time | Space | Bottleneck |
|------|------|-------|------------|
| Formula extraction | O(n) | O(1) streaming | XML parsing |
| Hardcoded vectors | O(n) | O(v) | XML streaming |
| Formula grouping | O(f) | O(f) | Grouping dicts |
| Named ranges | O(d + s) | O(d) | openpyxl reads |
| Transformation building | O(f × r_avg) | O(v) | Overlap checks |
| Edge dedup | O(e) | O(e) | Set operations |
| Upstream tracing | O(u × m × L) | O(u × L) | Matrix multiply |
| Formula→SQL | O(f × len) | O(len) | Regex + parsing |
| LLM naming | O(v / batch) | O(batch) | API latency |
| Contract writing | O(v + t + e) | O(v + t + e) | Workbook I/O |
| Mermaid generation | O(v + e) | O(v + e) | String concat |
| Python generation | O(v + t) | O(v + t) | Code generation |
| Topo sort | O(t + e) | O(t) | DFS |

**Where:**
- n = total cells in workbook
- f = formula cells
- v = variables
- t = transformations
- e = dependency edges
- d = defined names
- s = sample value reads
- u = upstream files
- m = model vectors
- L = average vector length
- r_avg = average variable ranges per sheet
- len = average formula string length
- batch = LLM batch size (default 20)

**Overall pipeline:** Dominated by upstream tracing (matrix operations) and
LLM calls (network latency). For a 100-sheet file with no upstream tracing
and skip_llm=True, the pipeline completes in < 1 second.
