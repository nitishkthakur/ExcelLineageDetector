# Excel Lineage Detector

You are an expert Python developer. Build a production-quality command-line tool that takes any Excel file as input and forensically extracts every data connection embedded within it — no matter where it's hidden.

## The Problem

Excel files are often connected to dozens of data sources: databases, APIs, upstream files, cloud services. These connections are scattered across formulas, VBA code, Power Query, pivot tables, named ranges, comments, hidden sheets, and more. Most people have no idea what their Excel file is actually connected to. This tool makes that visible.

## What to Build

A single command does everything:
```bash
python detect_lineage.py path/to/file.xlsx
```

It produces three outputs, all named after the input file:
- **JSON** — full structured extraction, machine-readable
- **Excel report** — human-readable lineage report for business users
- **PNG graph** — visual map of data flow into the file

## Detection Requirements

Think like a forensic investigator. Scan **every** place a data connection could hide in an Excel file — cells, formulas, VBA, Power Query, pivot tables, external connections, named ranges, comments, hidden sheets, document properties, linked objects, and anything else you can think of. If there's a connection in there, find it.

For each connection found, capture everything you can: the source, query text, location in the file, author, timestamps, and any other available metadata. For SQL and other queries, parse them to extract the tables, columns, joins, and filters.

## Output Quality

- The JSON should be clean, hierarchical, and complete enough to drive downstream automation
- The Excel report should be usable by a non-technical analyst with no explanation — clear layout, good formatting, easy to navigate
- The graph should make the data flow immediately obvious — labeled, color-coded by source type, high resolution

## Implementation

Use whatever libraries best fit the job. Structure the code cleanly. Never crash on a single bad component — log and continue. Prioritize thoroughness over speed.

## Testing

Write self-validating tests. Programmatically generate tricky Excel test files that hide connection information in as many locations as possible, then prove the tool finds everything. Tests should produce a clear coverage score.