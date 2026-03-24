#!/usr/bin/env python3
"""Upstream Tracing — find the source of hardcoded vectors in a model file.

Given a model Excel file and a set of upstream Excel files, identifies which
upstream file / sheet / cell range each hardcoded vector in the model was
likely copy-pasted from.
"""

import argparse
import sys
from pathlib import Path

from lineage.tracing.config import TraceConfig
from lineage.tracing.tracer import UpstreamTracer
from lineage.tracing.report import TracingReporter


def main():
    parser = argparse.ArgumentParser(
        description="Trace hardcoded vectors in a model file back to upstream sources",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Trace sheet "Revenue" against two upstream files
  python trace_upstream.py model.xlsx --sheet "Revenue" \\
      --upstream source1.xlsx source2.xlsx

  # Trace against all xlsx files in a directory
  python trace_upstream.py model.xlsx --sheet "Revenue" \\
      --upstream-dir ./upstream_files/

  # Exact matching only (no approximate), custom config
  python trace_upstream.py model.xlsx --sheet "Revenue" \\
      --upstream source.xlsx --config my_config.json

  # List sheets in a file
  python trace_upstream.py model.xlsx --list-sheets
        """,
    )
    parser.add_argument("model_file", help="Path to the model Excel file")
    parser.add_argument(
        "--sheet", "-s",
        help="Name of the sheet to trace (required unless --list-sheets)",
    )
    parser.add_argument(
        "--upstream", "-u",
        nargs="+",
        help="One or more upstream Excel files",
    )
    parser.add_argument(
        "--upstream-dir",
        help="Directory containing upstream Excel files (.xlsx/.xlsm)",
    )
    parser.add_argument(
        "--config", "-c",
        help="Path to config file (JSON or YAML). Default: tracing_config.json if present",
    )
    parser.add_argument(
        "--out-dir", "-o",
        help="Output directory (default: same as model file)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Verbose logging",
    )
    parser.add_argument(
        "--list-sheets",
        action="store_true",
        help="List sheet names in the model file and exit",
    )
    parser.add_argument(
        "--max-level",
        type=int,
        default=10,
        help="Maximum recursion depth for formula-based tracing (default: 10)",
    )
    parser.add_argument(
        "--no-formula-tracing",
        action="store_true",
        help="Disable recursive formula-based external reference tracing",
    )
    args = parser.parse_args()

    model_path = Path(args.model_file).resolve()
    if not model_path.exists():
        print(f"Error: File not found: {model_path}", file=sys.stderr)
        sys.exit(1)

    # --list-sheets
    if args.list_sheets:
        from lineage.tracing.scanner import get_sheet_names
        names = get_sheet_names(model_path)
        if names:
            print(f"Sheets in {model_path.name}:")
            for i, n in enumerate(names, 1):
                print(f"  {i}. {n}")
        else:
            print("Could not read sheet names (is this a valid .xlsx/.xlsm file?)")
        sys.exit(0)

    # --sheet is required for tracing
    if not args.sheet:
        print("Error: --sheet is required. Use --list-sheets to see available sheets.",
              file=sys.stderr)
        sys.exit(1)

    # Validate sheet exists
    from lineage.tracing.scanner import get_sheet_names
    available = get_sheet_names(model_path)
    if available and args.sheet not in available:
        print(f"Error: Sheet '{args.sheet}' not found. Available sheets:", file=sys.stderr)
        for n in available:
            print(f"  - {n}", file=sys.stderr)
        sys.exit(1)

    # Collect upstream files
    upstream_paths: list[Path] = []
    if args.upstream:
        for f in args.upstream:
            p = Path(f).resolve()
            if not p.exists():
                print(f"Warning: Upstream file not found, skipping: {p}", file=sys.stderr)
                continue
            upstream_paths.append(p)
    if args.upstream_dir:
        d = Path(args.upstream_dir).resolve()
        if d.is_dir():
            for p in sorted(d.iterdir()):
                if p.suffix.lower() in (".xlsx", ".xlsm") and p != model_path:
                    upstream_paths.append(p)
        else:
            print(f"Warning: Not a directory: {d}", file=sys.stderr)

    if not upstream_paths:
        print("Error: No upstream files specified. Use --upstream or --upstream-dir.",
              file=sys.stderr)
        sys.exit(1)

    # Remove duplicates preserving order
    seen: set[Path] = set()
    deduped: list[Path] = []
    for p in upstream_paths:
        if p not in seen:
            seen.add(p)
            deduped.append(p)
    upstream_paths = deduped

    # Load config
    config = TraceConfig()
    config_path = Path(args.config) if args.config else Path("tracing_config.json")
    if config_path.exists():
        try:
            config = TraceConfig.from_file(config_path)
            if args.verbose:
                print(f"Loaded config from {config_path}")
        except Exception as e:
            print(f"Warning: Could not load config {config_path}: {e}", file=sys.stderr)

    # Output directory
    out_dir = Path(args.out_dir).resolve() if args.out_dir else model_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # Run tracing
    print(f"Model:    {model_path.name} -> sheet '{args.sheet}'")
    print(f"Upstream: {len(upstream_paths)} file(s)")
    if args.verbose:
        for p in upstream_paths:
            print(f"  - {p.name}")

    tracer = UpstreamTracer(config=config, verbose=args.verbose)
    matches, unmatched = tracer.trace(model_path, args.sheet, upstream_paths)

    # Summary
    n_model = len(set(m.model_range for m in matches)) + len(unmatched)
    n_exact = sum(1 for m in matches if m.match_type.startswith("exact"))
    n_approx = sum(1 for m in matches if m.match_type == "approximate")

    print(f"\nResults: {n_model} model vectors")
    print(f"  Exact matches:       {n_exact}")
    print(f"  Approximate matches: {n_approx}")
    print(f"  Unmatched vectors:   {len(unmatched)}")

    # ── Formula-based tracing ─────────────────────────────────────────
    level_refs = None
    if not args.no_formula_tracing:
        from lineage.tracing.formula_tracer import trace_formula_levels

        search_dirs = [model_path.parent]
        if args.upstream_dir:
            ud = Path(args.upstream_dir).resolve()
            if ud.is_dir() and ud != model_path.parent:
                search_dirs.append(ud)
        level_refs = trace_formula_levels(
            model_path,
            search_dirs=search_dirs,
            max_level=args.max_level,
            verbose=args.verbose,
        )
        if level_refs:
            total_refs = sum(len(r) for r in level_refs.values())
            total_found = sum(
                1 for refs in level_refs.values() for r in refs if r.file_found
            )
            total_missing = total_refs - total_found
            print(f"\nFormula tracing: {len(level_refs)} level(s), "
                  f"{total_refs} external ref(s) "
                  f"({total_found} found, {total_missing} missing)")
        else:
            print("\nFormula tracing: no external formula references found")

    # Write report
    reporter = TracingReporter()
    out_path = reporter.write_with_levels(
        matches, unmatched, config, model_path, args.sheet,
        upstream_paths, out_dir, level_refs=level_refs,
    )
    print(f"\nReport: {out_path}")


if __name__ == "__main__":
    main()
