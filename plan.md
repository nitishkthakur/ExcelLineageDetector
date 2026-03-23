# Excel Lineage Detector — Implementation Plan

## Goal

A single CLI tool (`detect_lineage.py`) that forensically extracts every data connection from any Excel file and produces three outputs: JSON, Excel report, PNG graph.

---

## 1. Detection Surface (What to Scan)

### 1.1 XML-Layer (XLSX = zip archive)
| Location | What it reveals |
|---|---|
| `xl/connections.xml` | ODBC/OLE DB/web/text data connections with full connection strings |
| `xl/queryTables/queryTable*.xml` | SQL queries, refresh settings, table bindings |
| `xl/pivotTables/pivotTable*.xml` | Pivot table data sources and cache references |
| `xl/pivotCache/pivotCacheDefinition*.xml` | Actual source for each pivot (sheet range, external file, OLAP) |
| `xl/workbook.xml` | Named ranges (including external cross-workbook ones), defined names |
| `xl/worksheets/sheet*.xml` | Cell formulas with external refs, hyperlinks, data validation |
| `xl/sharedStrings.xml` | URLs/paths embedded in string values |
| `xl/customXml/` | Power Query metadata (item0.xml, etc.) |
| `xl/vbaProject.bin` | VBA source code (compiled OLE binary) |
| `docProps/core.xml` | Author, company, created/modified timestamps |
| `docProps/app.xml` | Application metadata, manager, company |
| `docProps/custom.xml` | Custom document properties (often contain source metadata) |
| `[Content_Types].xml` | Reveals what components are present |

### 1.2 Formula-Layer (cell-level)
- **Cross-workbook refs**: `='C:\data\[source.xlsx]Sheet1'!$A$1`
- **UNC paths**: `='\\server\share\[file.xlsx]Sheet'!A1`
- **WEBSERVICE()**: pulls live data from a URL
- **RTD()**: Real-Time Data server connections
- **SQL.REQUEST()** (legacy add-in)
- **External named range refs**: `=ExternalName!$A$1`

### 1.3 VBA-Layer
Extract from `xl/vbaProject.bin` using `olevba`:
- `ADODB.Connection` / `ADODB.Recordset` — database connections
- `Connection.Open "..."` — connection strings
- `CommandText = "SELECT ..."` — embedded SQL
- `Workbooks.Open "..."` — file references
- `Shell`, `CreateObject("WScript.Shell")` — external process calls
- `XMLHTTP`, `WinHttp` — HTTP calls
- Regex patterns for: URLs, file paths, connection strings, SQL

### 1.4 Power Query (M Code)
Located in `xl/customXml/` and embedded in connections:
- `Source = ...` lines revealing data source type and address
- `Sql.Database(server, db)` — SQL Server
- `Oracle.Database(server)` — Oracle
- `OData.Feed(url)` — OData APIs
- `Web.Contents(url)` — generic web
- `File.Contents(path)` / `Excel.Workbook(File.Contents(...))` — file sources
- `SharePoint.Files(url)` — SharePoint
- `AzureStorage.Blobs(account)` — Azure

### 1.5 Pivot Tables
- Source range (internal sheet range)
- External data source (OLAP cube, external file)
- Cache definition pointing to a QueryTable (which has a SQL connection)

### 1.6 Hyperlinks
Scan `xl/worksheets/sheet*.xml` for `<hyperlink>` elements referencing external URLs and file paths.

### 1.7 Linked/Embedded OLE Objects
- Linked OLE objects point to external files — captured from `xl/drawings/` and relationship files.

### 1.8 Hidden Sheets
- Check `sheet state="hidden"` or `state="veryHidden"` in `workbook.xml`; scan them the same as visible sheets.

### 1.9 Comments / Notes
- Scan all comments (legacy notes and threaded comments) for URLs, file paths, server names.

---

## 2. Data Models

```python
@dataclass
class ParsedQuery:
    tables: list[str]
    columns: list[str]
    joins: list[dict]   # {"type": "INNER", "table": "orders", "on": "..."}
    filters: list[str]
    raw_sql: str

@dataclass
class DataConnection:
    id: str                         # deterministic hash
    category: str                   # "database" | "file" | "web" | "powerquery" | "vba" | "pivot" | "formula" | "hyperlink" | "ole" | "metadata"
    sub_type: str                   # "odbc" | "oledb" | "sql_server" | "oracle" | "xlsx" | "csv" | "odata" | ...
    source: str                     # human-readable source label
    raw_connection: str             # full connection string / URL / path
    location: str                   # "Sheet1!A1" | "VBA:Module1:line42" | "connections.xml" | ...
    query_text: str | None          # raw SQL or M code
    parsed_query: ParsedQuery | None
    author: str | None
    created_at: str | None
    modified_at: str | None
    metadata: dict                  # everything else
    confidence: float               # 0.0–1.0 (1.0 = definitive, <1.0 = heuristic)
```

---

## 3. Architecture

```
detect_lineage.py          ← CLI entry point (argparse)
lineage/
  __init__.py
  detector.py              ← ExcelLineageDetector orchestrator
  models.py                ← DataConnection, ParsedQuery dataclasses
  utils.py                 ← logging, hashing, path helpers

  extractors/
    base.py                ← BaseExtractor ABC
    connections.py         ← xl/connections.xml
    powerquery.py          ← Power Query M code
    formulas.py            ← Cell formula external refs
    externallinks.py       ← xl/externalLinks/ (resolves paths/URLs)
    vba.py                 ← olevba VBA extraction
    pivot.py               ← Pivot tables & cache
    querytable.py          ← Query tables
    hyperlinks.py          ← Hyperlinks
    namedranges.py         ← Named ranges with external refs
    comments.py            ← Comments / notes
    metadata.py            ← docProps/*
    ole.py                 ← Linked OLE objects
    hardcoded_values.py    ← Cells with value but no formula

  parsers/
    sql_parser.py          ← sqlglot: extract tables, columns, joins, filters
    m_parser.py            ← regex + heuristics for M code sources
    formula_parser.py      ← parse external formula refs
    connection_string.py   ← parse ODBC/OLE DB connection strings

  reporters/
    json_reporter.py       ← full structured JSON
    excel_reporter.py      ← formatted Excel workbook (openpyxl)
    graph_reporter.py      ← networkx + matplotlib PNG

  hardcoded_scanner.py     ← fast streaming vector scanner (lxml iterparse)

  tracing/                 ← upstream tracing module
    config.py              ← TraceConfig dataclass + JSON/YAML loader
    models.py              ← TracingVector, VectorMatch dataclasses
    scanner.py             ← streaming XML parsers (model + upstream)
    exact_matcher.py       ← hash-based lookup + batched numpy subsequence
    approx_matcher.py      ← vectorized similarity (Pearson/cosine/Euclidean)
    tracer.py              ← orchestrator with parallel upstream scanning
    formula_tracer.py      ← recursive external formula reference tracing
    report.py              ← Excel report: Config + Tracing Results + Level N sheets

trace_upstream.py          ← CLI entry point for upstream tracing
tracing_config.json        ← default config for upstream tracing

tests/
  test_generator.py        ← programmatically build tricky Excel test files
  test_detector.py         ← 19 tests: detection coverage and all reporters
  test_tracing.py          ← 34 tests: upstream tracing and formula tracing
  fixtures/                ← generated .xlsx test files (gitignored)
```

---

## 4. Extraction Logic Details

### connections.py
Parse `xl/connections.xml` with lxml. Each `<connection>` has:
- `type` attribute (1=ODBC, 2=DAO, 3=web, 4=OLE DB, 5=text, 6=ADO, 7=DSP)
- `dbPr` child → `connection` attribute (connection string)
- `webPr` child → `url` attribute
- `olapPr` child → `localConnection`
- `textPr` child → `sourceFile`
Map to `DataConnection` with full connection string parsing.

### powerquery.py
Power Query lives in two places:
1. `xl/customXml/item0.xml` — JSON blob with all queries
2. Embedded M code in connections named `Query - X`

Parse the JSON blob; for each query, extract M formula and run through `m_parser.py`.

### formulas.py
Use `openpyxl` to iterate every cell in every sheet (including hidden). Regex-match:
- `\[([^\]]+\.xls[xmb]?)\]` → external workbook
- `'([A-Za-z]:\\[^']+)\[` or `'(\\\\[^']+)\[` → UNC/local paths
- `WEBSERVICE\(([^)]+)\)` → web service calls
- `RTD\(` → RTD connections

### vba.py
Use `olevba` (`oletools.olevba`) to decompile `xl/vbaProject.bin`. Parse source:
- Regex for connection strings: `(Provider|Data Source|Server|Database|DSN)\s*=\s*[^;"`]+`
- Regex for SQL: `(SELECT|INSERT|UPDATE|DELETE|EXEC)\s+.+` (multiline, case-insensitive)
- Regex for file paths: `[A-Za-z]:\\[^\s"']+` and `\\\\[^\s"']+`
- Regex for URLs: `https?://[^\s"']+`

### sql_parser.py
Use `sqlglot` to parse extracted SQL:
```python
ast = sqlglot.parse_one(sql)
tables = [t.name for t in ast.find_all(sqlglot.exp.Table)]
columns = [c.name for c in ast.find_all(sqlglot.exp.Column)]
```
Fallback to regex if sqlglot fails to parse.

---

## 5. Reporter Details

### JSON (`{filename}_lineage.json`)
```json
{
  "file": "report.xlsx",
  "scanned_at": "2024-01-15T10:30:00Z",
  "summary": {
    "total_connections": 14,
    "by_category": {"database": 3, "powerquery": 4, ...}
  },
  "connections": [ ... DataConnection objects ... ]
}
```

### Excel Report (`{filename}_lineage_report.xlsx`)
Sheet 1 — **All Connections**: full table with auto-filters, frozen headers, color-coded rows by category
Sheets 2-N — **Per-sheet hardcoded vector sheets**: one sheet per original workbook sheet showing hardcoded numeric vectors (column/row direction, cell range, length, sample values)

### Graph (`{filename}_lineage.png`)
- **Layout**: left-to-right hierarchical (sources → center node = target file)
- **Nodes**: target file (large, center); each unique source (sized by connection count)
- **Color by category**:
  - `#2196F3` blue = database
  - `#4CAF50` green = file
  - `#FF9800` orange = web/API
  - `#9C27B0` purple = Power Query
  - `#F44336` red = VBA
  - `#607D8B` gray = pivot/other
- **Edge labels**: connection sub-type
- **Legend** included
- **Resolution**: 300 DPI, minimum 2400×1600px

---

## 6. Testing Strategy

### Test File Generator (`test_generator.py`)
Programmatically create `test_connections.xlsx` with planted connections:

| # | Type | Method |
|---|---|---|
| 1 | Cross-workbook formula | `='[source.xlsx]Sheet1'!A1` in cell A1 |
| 2 | UNC path formula | `='\\server\share\[data.xlsx]Sheet1'!A1` |
| 3 | WEBSERVICE formula | `=WEBSERVICE("https://api.example.com/data")` |
| 4 | External named range | Named range pointing to external workbook |
| 5 | ODBC connection | Add to `xl/connections.xml` via XML manipulation |
| 6 | OLE DB connection | Add to `xl/connections.xml` |
| 7 | Web query | `<webPr url="https://data.example.com/table">` |
| 8 | Power Query (file) | M code: `Excel.Workbook(File.Contents("C:\data.xlsx"))` |
| 9 | Power Query (SQL) | M code: `Sql.Database("server", "db")` |
| 10 | Power Query (web) | M code: `Web.Contents("https://api.example.com")` |
| 11 | VBA ADODB | VBA module with `ADODB.Connection` + SQL string |
| 12 | VBA file open | VBA with `Workbooks.Open "C:\source.xlsx"` |
| 13 | VBA HTTP | VBA with `XMLHTTP.Open "GET", "https://api.example.com"` |
| 14 | Pivot external | Pivot table pointing to external data source |
| 15 | Hyperlink external | Hyperlink to external file or URL |
| 16 | Comment with URL | Comment containing `https://datasource.example.com` |
| 17 | Hidden sheet formula | External ref on a hidden sheet |
| 18 | Custom property | docProps/custom.xml with source metadata |
| 19 | Query table SQL | `<queryTable>` with embedded SQL |
| 20 | OLE linked object | Linked OLE object reference |

### Coverage Score
```
coverage_score = len(found_ids & planted_ids) / len(planted_ids) * 100
```
Tests pass if score ≥ 60%. Current coverage: **100%**. Prints per-type breakdown.

---

## 7. Dependencies

```
openpyxl>=3.1        # Excel read/write
lxml>=4.9            # fast XML parsing
oletools>=0.60       # VBA extraction (olevba)
sqlglot>=18.0        # SQL parsing
networkx>=3.0        # graph structure
matplotlib>=3.7      # graph rendering
Pillow>=10.0         # image handling
```

Optional (graceful degradation if absent):
```
xlrd>=2.0            # .xls support
python-pptx          # if we ever need shape-linked sources
```

---

## 8. Error Handling Philosophy

- Each extractor runs in a `try/except`; failures are logged and skipped — never crash the whole run
- If `olevba` fails (no VBA), log at DEBUG level (not an error)
- If `sqlglot` fails to parse SQL, store raw text and skip parsed_query
- Missing XML files in the zip are silently skipped (not all Excel files have all parts)
- Invalid XML is caught and logged

---

## 9. CLI Interface

```bash
python detect_lineage.py file.xlsx              # standard run
python detect_lineage.py file.xlsx --verbose    # debug logging
python detect_lineage.py file.xlsx --json-only  # skip Excel/graph outputs
python detect_lineage.py file.xlsx --out-dir /tmp/reports
```

Output files:
```
file_lineage.json
file_lineage.xlsx
file_lineage.png
```

---

## 10. File Count & Scope

| File | Purpose |
|---|---|
| `detect_lineage.py` | CLI entry + wiring |
| `lineage/detector.py` | Orchestrator |
| `lineage/models.py` | Data classes |
| `lineage/utils.py` | Shared helpers |
| `lineage/extractors/*.py` | 13 extractor modules |
| `lineage/parsers/*.py` | 4 parser modules |
| `lineage/reporters/*.py` | 3 reporter modules |
| `tests/test_generator.py` | Builds test Excel files |
| `tests/test_detector.py` | 19 tests — validates detection coverage and reporters |
| `tests/test_tracing.py` | 34 tests — upstream tracing and formula tracing |
| `trace_upstream.py` | CLI for upstream tracing |
| `tracing_config.json` | Default upstream tracing config |
| `lineage/tracing/*.py` | 8 tracing modules |
| `requirements.txt` | Dependencies |

---

## Approval Checklist

- [ ] Detection surface is exhaustive enough
- [ ] Data model fields are sufficient
- [ ] Graph layout/style is acceptable
- [ ] Excel report sheet structure looks right
- [ ] Testing coverage threshold (80%) is appropriate
- [ ] Any additional connection types to add?
