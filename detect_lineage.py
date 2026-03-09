#!/usr/bin/env python3
"""Excel Lineage Detector - CLI entry point.

Forensically extracts data connections from Excel files (.xlsx, .xlsm, .xlsb).
"""

import argparse
import sys
from pathlib import Path

from lineage.detector import ExcelLineageDetector
from lineage.reporters.json_reporter import JsonReporter
from lineage.reporters.excel_reporter import ExcelReporter
from lineage.reporters.graph_reporter import GraphReporter
from lineage.utils import set_log_level


def main():
    parser = argparse.ArgumentParser(
        description="Forensically extract data connections from Excel files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python detect_lineage.py workbook.xlsx
  python detect_lineage.py workbook.xlsm --verbose
  python detect_lineage.py workbook.xlsx --json-only --out-dir ./results
        """,
    )
    parser.add_argument(
        "file",
        help="Path to Excel file (.xlsx, .xlsm, .xlsb)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose/debug logging",
    )
    parser.add_argument(
        "--json-only",
        action="store_true",
        help="Skip Excel report and graph outputs, only write JSON",
    )
    parser.add_argument(
        "--out-dir",
        default=None,
        help="Output directory (default: same directory as input file)",
    )
    args = parser.parse_args()

    set_log_level(args.verbose)

    input_path = Path(args.file).resolve()
    if not input_path.exists():
        print(f"Error: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if not input_path.is_file():
        print(f"Error: Not a file: {input_path}", file=sys.stderr)
        sys.exit(1)

    suffix = input_path.suffix.lower()
    if suffix not in (".xlsx", ".xlsm", ".xlsb", ".xls"):
        print(f"Warning: Unexpected file extension: {suffix}", file=sys.stderr)

    out_dir = Path(args.out_dir).resolve() if args.out_dir else input_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"Scanning: {input_path}")
    detector = ExcelLineageDetector()
    connections = detector.detect(input_path)
    print(f"Found {len(connections)} connection(s)")

    if connections:
        # Show quick summary
        from collections import Counter
        by_cat = Counter(c.category for c in connections)
        for cat, count in sorted(by_cat.items(), key=lambda x: -x[1]):
            print(f"  {cat}: {count}")

    # Write JSON report (always)
    json_reporter = JsonReporter()
    json_path = json_reporter.write(connections, input_path, out_dir)
    print(f"JSON: {json_path}")

    if not args.json_only:
        # Write Excel report
        try:
            excel_reporter = ExcelReporter()
            xl_path = excel_reporter.write(connections, input_path, out_dir)
            print(f"Excel: {xl_path}")
        except Exception as e:
            print(f"Warning: Excel report failed: {e}", file=sys.stderr)

        # Write graph
        try:
            graph_reporter = GraphReporter()
            png_path = graph_reporter.write(connections, input_path, out_dir)
            print(f"Graph: {png_path}")
        except Exception as e:
            print(f"Warning: Graph report failed: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
