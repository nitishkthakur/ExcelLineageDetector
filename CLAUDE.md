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
   - If it needs a dedicated Excel sheet, add a `_write_*_sheet` method and register it in `write()` and add the category to `DEDICATED_SHEET_CATEGORIES`; otherwise it auto-appears in the "Other Sources" sheet
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

**Hardcoded / copy-pasted values (`input` category)**: `HardcodedValuesExtractor` detects cells that have a value but no formula. It reports them at four confidence levels — `named_input` (named range → single cell, conf 0.95), `hardcoded_value` (labeled row in an Inputs/Assumptions sheet, conf 0.7–0.9), `pasted_table` (contiguous 2×2+ numeric block with ≥2 column headers, conf 0.75), and `source_note` (text cell matching "Source: Bloomberg", "Per FactSet", etc., conf 0.7). Cells with no label context are skipped to suppress noise. The Excel report writes these to a dedicated "Hardcoded Inputs" sheet with a **Sheet** column indicating which worksheet they came from.

**Finance terminal formulas**: `FormulasExtractor` detects proprietary add-in functions that pull live data: Bloomberg (`BDP`, `BDH`, `BDS`, `BQL`), Reuters/Refinitiv (`RHistory`, `RData`, `TR.*`), FactSet (`FDS`, `FQL`), Capital IQ (`CIQ`, `CIQCONTENT`), SNL (`SNLD`, `SNLC`), and Wind Info (`WSD`, `WSS`). Sub-types match the terminal name (e.g. `bloomberg`, `reuters`, `factset`).

**XML namespace handling**: All extractors try three forms for every element lookup — with namespace `{http://schemas.openxmlformats.org/spreadsheetml/2006/main}tag`, wildcard `{*}tag`, and bare `tag` — because namespace presence varies across Excel versions.

## Tests

`tests/test_generator.py` builds a test `.xlsx` programmatically with planted connections of every type. `tests/test_detector.py::test_coverage` runs the detector against it and asserts ≥60% detection rate. Re-run the generator if you add new connection types to test:

```bash
.venv/bin/python tests/test_generator.py
```

There are 19 tests in `test_detector.py` covering: coverage rate, required fields, deduplication, serialisation, all three reporters (JSON/Excel/PNG), all four parsers, and specific assertions for hardcoded values, source notes, Bloomberg formulas, SharePoint external links, SQL table extraction, and the "Hardcoded Inputs" Excel sheet.

**When adding a new planted connection type** to the generator: add it to the `planted` list at the bottom of `generate_test_workbook()`, inject its XML/data in one of the `Step 3` blocks, and add a focused test function in `test_detector.py`.
