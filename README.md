# ExcelLineageDetector

Forensically extract every data connection hidden inside an Excel file — databases, Power Query, VBA, formulas, pivot tables, hyperlinks, comments, and more — and produce a JSON report, a formatted Excel workbook, and a visual lineage graph.

## Outputs

For an input file `report.xlsx`, the tool writes three files:

| File | Description |
|---|---|
| `report_lineage.json` | Full structured extraction, machine-readable |
| `report_lineage_report.xlsx` | Human-readable report with 7 categorized sheets |
| `report_lineage_graph.png` | Hierarchical visual map of data flow into the file |

## Installation

**Python 3.10+ required.**

```bash
# Clone the repo
git clone https://github.com/your-org/ExcelLineageDetector.git
cd ExcelLineageDetector

# Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate        # Linux / macOS
# .venv\Scripts\activate         # Windows

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
python detect_lineage.py path/to/file.xlsx
```

All three output files are written to the same directory as the input file by default.

### Options

```
python detect_lineage.py <file> [options]

positional arguments:
  file              Path to Excel file (.xlsx, .xlsm, .xlsb)

options:
  --out-dir DIR     Write outputs to DIR instead of the input file's directory
  --json-only       Skip the Excel report and PNG graph (faster)
  --verbose, -v     Enable debug logging
```

### Examples

```bash
# Basic run — outputs written next to the input file
python detect_lineage.py reports/Q4_dashboard.xlsx

# Write outputs to a specific folder
python detect_lineage.py reports/Q4_dashboard.xlsx --out-dir ./lineage_out

# JSON only (no Excel report or graph)
python detect_lineage.py reports/Q4_dashboard.xlsx --json-only

# Verbose logging to see what each extractor finds
python detect_lineage.py reports/Q4_dashboard.xlsx --verbose
```

## What Gets Detected

| Source | Where it hides |
|---|---|
| Database connections | `xl/connections.xml` (ODBC, OLE DB, OLAP) |
| Power Query (M code) | `xl/customXml/item0.xml`, connection names |
| Cross-workbook formulas | Cell formulas: `='[source.xlsx]Sheet1'!A1` |
| UNC / local file paths | Cell formulas with `\\server\share\[file.xlsx]` |
| WEBSERVICE / RTD | Live-data formula functions |
| VBA (ADODB, SQL, HTTP) | `xl/vbaProject.bin` decompiled with olevba |
| Pivot table sources | `xl/pivotCache/pivotCacheDefinition*.xml` |
| Query tables | `xl/queryTables/` with embedded SQL |
| Hyperlinks | External URLs and file paths in cells |
| Named ranges | External workbook references in defined names |
| Comments | URLs and file paths embedded in cell notes |
| Document properties | `docProps/core.xml`, `app.xml`, `custom.xml` |
| Linked OLE objects | External file links in drawings/relationships |
| Hidden sheets | Same extraction, regardless of sheet visibility |

## Running the Tests

```bash
# Run all tests
python -m pytest tests/ -v

# Run only the coverage test (plants 20+ connection types, scores detection rate)
python -m pytest tests/test_detector.py::test_coverage -v -s
```

The coverage test programmatically generates a tricky Excel file with planted connections across every supported category and verifies the detector finds at least 60% of them. Current coverage: **100%**.

## Project Structure

```
detect_lineage.py          CLI entry point
requirements.txt
lineage/
  detector.py              Orchestrator — runs all extractors, deduplicates
  models.py                DataConnection and ParsedQuery dataclasses
  utils.py                 Logging helpers
  extractors/              One module per connection source type (11 total)
  parsers/                 SQL (sqlglot), Power Query M, formula ref, connection string
  reporters/
    json_reporter.py       Writes _lineage.json
    excel_reporter.py      Writes _lineage_report.xlsx (7 sheets)
    graph_reporter.py      Writes _lineage_graph.png
tests/
  test_generator.py        Programmatically builds a test .xlsx with planted connections
  test_detector.py         Validates detection coverage and all reporters
  fixtures/                Generated test files (git-ignored)
```

## Sample Graph Output

The PNG graph groups sources by category on the left, with all connections flowing through a bus line to the target Excel file on the right.

![Sample lineage graph](sample_output/debug_test_lineage_graph.png)
