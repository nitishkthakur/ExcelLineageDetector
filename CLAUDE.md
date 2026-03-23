# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

Always use the project venv at `.venv/`:

```bash
# Run all tests
.venv/bin/python -m pytest tests/ -v

# Run a single test
.venv/bin/python -m pytest tests/test_detector.py::test_coverage -v -s

# Run the detector on a file
.venv/bin/python detect_lineage.py path/to/file.xlsx --verbose

# JSON output only (skips Excel report and PNG graph)
.venv/bin/python detect_lineage.py path/to/file.xlsx --json-only --out-dir ./out

# Upstream tracing — trace hardcoded vectors back to source files
.venv/bin/python trace_upstream.py model.xlsx --sheet "Sheet1" --upstream-dir ./sources/
.venv/bin/python trace_upstream.py model.xlsx --list-sheets

# Formula tracing options
.venv/bin/python trace_upstream.py model.xlsx --sheet "Sheet1" --upstream-dir ./sources/ --max-level 3
.venv/bin/python trace_upstream.py model.xlsx --sheet "Sheet1" --upstream-dir ./sources/ --no-formula-tracing

# Convert upstream tracing report to Mermaid flowchart
.venv/bin/python trace_upstream_mermaid.py upstream_tracing_model.xlsx
.venv/bin/python trace_upstream_mermaid.py upstream_tracing_model.xlsx --lr -o diagram.md
```

## Architecture

An Excel file (`.xlsx`/`.xlsm`/`.xlsb`) is a ZIP archive. The detector opens it as both a `zipfile.ZipFile` (for raw XML access) and an `openpyxl` workbook (for cell-level access), then runs all 13 extractors in sequence.

**Data flow:**
```
detect_lineage.py (CLI)
  → ExcelLineageDetector.detect()       # lineage/detector.py
      → each Extractor.extract(zip, wb) # lineage/extractors/*.py
          → Parsers                     # lineage/parsers/*.py
      → deduplicate by DataConnection.id
  → Reporters                           # lineage/reporters/*.py
```

**`DataConnection`** (`lineage/models.py`) is the single output type. Key fields: `category` (database|file|web|powerquery|vba|pivot|hyperlink|ole|metadata|formula|**input**), `sub_type` (more granular), `raw_connection` (full path/URL/string), `location` (where found, e.g. `Sheet1!A10` or `VBA:Module1:42`), `confidence` (0.0–1.0), `parsed_query` (populated for database connections with SQL command text).

IDs are `sha256(category|raw_connection|location)[:12]` — deduplication is by this ID, first occurrence wins.

## Extractors

Each extractor in `lineage/extractors/` handles one ZIP region. Never raise exceptions — catch and log, return partial results.

**Checklist when adding a new extractor:**
1. Create class inheriting `BaseExtractor`, implement `extract(zip_file, workbook) -> list[DataConnection]`
2. Register in `EXTRACTORS` list in `lineage/detector.py`
3. If the extractor introduces a **new `category`** string (beyond the 10 existing ones):
   - Add a color entry to `CATEGORY_COLORS` in `lineage/reporters/excel_reporter.py`
   - Add color entries to `CATEGORY_COLORS` and `CATEGORY_BG` in `lineage/reporters/graph_reporter.py`
   - Add the category name to `CATEGORY_ORDER` in `graph_reporter.py`
   - The Excel reporter no longer has dedicated per-category sheets; all connections go to "All Connections". No reporter changes needed for new categories (CATEGORY_COLORS still applies for row colour coding)
4. JSON output is fully automatic — no changes needed

| Extractor | ZIP path(s) it reads |
|---|---|
| `ConnectionsExtractor` | `xl/connections.xml` |
| `PowerQueryExtractor` | `xl/customXml/*.xml`, `xl/connections.xml` (Query- prefix) |
| `FormulasExtractor` | `xl/worksheets/sheet*.xml` — cell `<f>`, `<dataValidation>`, `<cfRule>` |
| `ExternalLinksExtractor` | `xl/externalLinks/externalLink*.xml` + their `.rels` files |
| `VbaExtractor` | `xl/vbaProject.bin` (via oletools olevba, raw binary fallback) |
| `PivotExtractor` | `xl/pivotTables/`, `xl/pivotCache/` |
| `QueryTableExtractor` | `xl/queryTables/` |
| `HyperlinksExtractor` | Sheet XML hyperlink elements + `xl/worksheets/_rels/` |
| `NamedRangesExtractor` | `xl/workbook.xml` definedName elements |
| `CommentsExtractor` | `xl/comments*.xml`, `xl/threadedComments/` |
| `MetadataExtractor` | `docProps/core.xml`, `app.xml`, `custom.xml` |
| `OleExtractor` | `xl/drawings/_rels/`, `xl/worksheets/_rels/` |
| `HardcodedValuesExtractor` | `xl/worksheets/sheet*.xml` — cells with value but no formula |

## Parsers

- **`m_parser.py`**: `parse_all(m_code)` returns *all* data sources found in a Power Query M script (one dict per source). `parse(m_code)` is a backward-compat wrapper returning only the first. Covers 40+ M functions (SQL Server, Snowflake, BigQuery, SharePoint, Salesforce, SAP, etc.).
- **`formula_parser.py`**: Parses Excel external reference syntax — handles local paths, UNC paths, and HTTPS/SharePoint URL prefixes before `[workbook.xlsx]`.
- **`connection_string.py`**: Parses ODBC/OLE DB connection strings into structured fields; identifies database type from Provider/Driver.
- **`sql_parser.py`**: Uses `sqlglot` for AST-based SQL parsing (tables, columns, joins); regex fallback. Called by `ConnectionsExtractor`, `PivotExtractor`, and `QueryTableExtractor` to populate `parsed_query` on database connections.

## Key Patterns

**External workbook resolution**: Formulas may contain either `'[budget.xlsx]Sheet1'!A1` (literal) or `=[1]Sheet1!A1` (numeric index). The numeric index `[n]` maps to `xl/externalLinks/externalLink{n}.xml`, where the `.rels` file holds the full resolved path (which may be a SharePoint URL like `https://company.sharepoint.com/.../budget.xlsx`). `ExternalLinksExtractor` handles this; `FormulasExtractor` detects both forms.

**Power Query multi-source queries**: A single M script can join multiple databases/files. `PowerQueryExtractor` calls `parse_m_all()` and emits one `DataConnection` per source found in the script.

**Hardcoded / copy-pasted values (`input` category)**: `HardcodedValuesExtractor` detects cells that have a value but no formula. It reports them at four confidence levels — `named_input` (named range → single cell, conf 0.95), `hardcoded_value` (labeled row in an Inputs/Assumptions sheet, conf 0.7–0.9), `pasted_table` (contiguous 2×2+ numeric block with ≥2 column headers, conf 0.75), and `source_note` (text cell matching "Source: Bloomberg", "Per FactSet", etc., conf 0.7). Cells with no label context are skipped to suppress noise.

**Hardcoded vector scanner** (`lineage/hardcoded_scanner.py`): Fast standalone scanner (separate from the extractor pipeline) used by `ExcelReporter` to detect **vectors** — contiguous runs of hardcoded (non-formula) numeric cells in a single row or column (minimum length 3). Uses `lxml.etree.iterparse` for O(n) streaming with constant memory — suitable for very large files. Key API: `scan_vectors(path) -> dict[sheet_name, list[HardcodedVector]]`. Returns `{}` for XLS (binary) files. Each `HardcodedVector` has `cell_range`, `direction` ("row"/"column"), `length`, `start_cell`, `end_cell`, `sample_values` (first 5).

**Finance terminal formulas**: `FormulasExtractor` detects proprietary add-in functions that pull live data: Bloomberg (`BDP`, `BDH`, `BDS`, `BQL`), Reuters/Refinitiv (`RHistory`, `RData`, `TR.*`), FactSet (`FDS`, `FQL`), Capital IQ (`CIQ`, `CIQCONTENT`), SNL (`SNLD`, `SNLC`), and Wind Info (`WSD`, `WSS`). Sub-types match the terminal name (e.g. `bloomberg`, `reuters`, `factset`).

**XML namespace handling**: All extractors try three forms for every element lookup — with namespace `{http://schemas.openxmlformats.org/spreadsheetml/2006/main}tag`, wildcard `{*}tag`, and bare `tag` — because namespace presence varies across Excel versions.

## Tests

`tests/test_generator.py` builds a test `.xlsx` programmatically with planted connections of every type. `tests/test_detector.py::test_coverage` runs the detector against it and asserts ≥60% detection rate. Re-run the generator if you add new connection types to test:

```bash
.venv/bin/python tests/test_generator.py
```

There are 72 tests total: 19 in `test_detector.py` (coverage rate, required fields, deduplication, serialisation, all three reporters, all four parsers, hardcoded values, source notes, Bloomberg formulas, SharePoint external links, SQL table extraction, Excel report structure) and 53 in `test_tracing.py` (config, scanner, exact/approximate matchers, batch kernels, end-to-end tracer, report, formula tracer helpers, streaming parser, cell filter, regex parsing, file resolution, multi-level tracing, report with Level sheets, precedent walking unit tests, transitive integration tests, precedent chain report rendering).

**When adding a new planted connection type** to the generator: add it to the `planted` list at the bottom of `generate_test_workbook()`, inject its XML/data in one of the `Step 3` blocks, and add a focused test function in `test_detector.py`.

## Upstream Tracing (`lineage/tracing/`)

A separate tool (`trace_upstream.py`) that identifies where hardcoded vectors in a model file were originally copy-pasted from, by matching against a set of upstream Excel files. See `upstream_algorithm.md` for the full algorithm description.

**Data flow:**
```
trace_upstream.py (CLI)
  → TraceConfig                        # lineage/tracing/config.py
  → UpstreamTracer.trace()             # lineage/tracing/tracer.py
      → scan_model_sheet()             # lineage/tracing/scanner.py (hardcoded only)
      → scan_upstream_file() × N       # parallel, ALL numerics (formula + hardcoded)
      → ExactMatcher.match()           # lineage/tracing/exact_matcher.py
      → ApproximateMatcher.match()     # lineage/tracing/approx_matcher.py
  → trace_formula_levels()             # lineage/tracing/formula_tracer.py (recursive)
  → TracingReporter.write_with_levels()# lineage/tracing/report.py

trace_upstream_mermaid.py (CLI — post-processing)
  → reads upstream_tracing_*.xlsx (Level N sheets)
  → builds Mermaid flowchart (file → sheet → range edges across levels)
  → writes *_mermaid.md
```

**Key modules:**

| Module | Purpose |
|--------|---------|
| `config.py` | `TraceConfig` dataclass + JSON/YAML loader |
| `models.py` | `TracingVector` (full values), `VectorMatch` (result) |
| `scanner.py` | Streaming XML parsers: `_stream_all_numerics` (upstream) and `_stream_hardcoded_numerics` (model). Reuses helpers from `hardcoded_scanner.py`. |
| `exact_matcher.py` | Hash-based O(1) lookup + batched numpy subsequence matching |
| `approx_matcher.py` | Vectorized numpy similarity (Pearson/cosine/Euclidean) with sliding window |
| `tracer.py` | Orchestrator: parallel scanning, match coordination, result assembly |
| `report.py` | Excel report writer: Config sheet + Tracing Results sheet + Level N sheets |
| `formula_tracer.py` | Recursive formula-based external reference tracing (Level 1, 2, ...) |

**Formula tracing** (`formula_tracer.py`): Scans formulas referencing external workbooks (`'[file.xlsx]Sheet'!A1` or `[1]Sheet!A1`), then recursively follows those references through multiple levels. Level 1 scans the entire model file; Level 2+ scans only the cell ranges identified at the previous level. Uses `CellFilter` for rectangle-based scoping and `_get_link_map()` to resolve numeric external link indices via `.rels` files. **Transitive precedent walking**: For Level 2+, if a target cell's formula does NOT directly reference an external file but depends on other cells that do (through arbitrary intermediate formulas), the in-workbook dependency graph is walked via BFS until external references are found or the chain dead-ends. Handles cross-sheet references, circular reference safety (`visited` set), and caps at `_MAX_PRECEDENT_DEPTH=20` / `_MAX_CELLS_VISITED=10,000`. The `ExternalReference.precedent_chain` field records the intermediate cells traversed. Stops when upstream file is not found on disk or `max_level` is reached. Cycle prevention via `visited_files` set.

**Configuration** (`tracing_config.json`):
- `matching.exact` / `matching.approximate` — enable/disable match modes
- `matching.top_n` — top-N approximate matches per vector (default 5)
- `matching.similarity_metric` — `pearson` (default), `cosine`, or `euclidean`
- `matching.min_similarity` — minimum similarity threshold (default 0.8)
- `matching.subsequence_matching` — match model as subsequence of longer upstream vector
- `matching.length_tolerance_pct` — allow ±N% length mismatch for approximate (default 50%)
- `matching.direction_sensitive` — if false (default), column↔row matching allowed
- `performance.max_workers` — null = auto (cpu_count)

**Performance design:**
- Upstream files scanned with `ProcessPoolExecutor` (parallel I/O + XML parsing)
- All numpy arrays pre-stacked during `index_upstream()` — one matrix per length bucket
- Exact: batched `sliding_window_view` on 2D arrays + vectorized `==` comparison
- Approximate: single matrix multiply per length bucket (batch Pearson correlation)
- Benchmark: 91 model vectors × 3196 upstream vectors = **~0.9s total** (including scan + match + report)
