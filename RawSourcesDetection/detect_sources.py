#!/usr/bin/env python3
"""detect_sources.py — CLI entry point for RawSourcesDetection.

Usage:
    # From RawSourcesDetection/ folder:
    python detect_sources.py model.xlsx --sheets "Sheet1,Inputs"
    python detect_sources.py model.xlsx --sheets "Inputs" --verbose
    python detect_sources.py path/to/model.xlsx --sheets "Sheet1" --inputs-dir ./inputs

    # Full options:
    python detect_sources.py model.xlsx \\
        --sheets "Sheet1,Assumptions" \\
        --inputs-dir ./inputs \\
        --config config.json \\
        --out-dir ./output \\
        --verbose

Output:
    RawSourcesDetection/output/raw_sources_<model_stem>.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Add project root (parent of RawSourcesDetection/) to sys.path
_THIS = Path(__file__).resolve().parent
sys.path.insert(0, str(_THIS.parent))

from pipeline.config import RSDConfig
from pipeline.orchestrator import run
from pipeline.report_writer import write_report


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Detect and document all upstream data sources of an Excel model. "
            "Traces formula-based external references recursively and matches "
            "hardcoded numeric vectors against known input files."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "model",
        help=(
            "Model Excel file. Can be a bare filename (looked up in model/ folder), "
            "or any absolute / relative path."
        ),
    )
    parser.add_argument(
        "--sheets", required=True,
        help='Comma-separated sheet names to trace, e.g. "Sheet1,Inputs,Assumptions"',
    )
    parser.add_argument(
        "--inputs-dir", default=None,
        help=(
            "Directory containing upstream input files. "
            "Searched recursively for .xlsx/.xlsm files. "
            "Default: RawSourcesDetection/inputs/"
        ),
    )
    parser.add_argument(
        "--config", default=None,
        help=(
            "Path to config.json. "
            "Default: RawSourcesDetection/config.json"
        ),
    )
    parser.add_argument(
        "--out-dir", default=None,
        help="Output directory for the Excel report. Default: RawSourcesDetection/output/",
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Print progress to stdout.",
    )
    parser.add_argument(
        "--max-levels", type=int, default=None,
        help="Override max_formula_levels from config (e.g. --max-levels 3).",
    )
    parser.add_argument(
        "--approximate", action="store_true", default=None,
        help="Enable approximate vector matching in addition to exact.",
    )

    args = parser.parse_args()

    rsd_root = _THIS

    # ── Resolve model path ───────────────────────────────────────────────────
    model_path = Path(args.model)
    if not model_path.is_absolute():
        # Try model/ subfolder first, then cwd
        candidate = rsd_root / "model" / args.model
        if candidate.exists():
            model_path = candidate
        else:
            model_path = model_path.resolve()

    if not model_path.exists():
        print(f"Error: model file not found: {model_path}", file=sys.stderr)
        sys.exit(1)

    # ── Resolve inputs dir ───────────────────────────────────────────────────
    inputs_dir = Path(args.inputs_dir) if args.inputs_dir else rsd_root / "inputs"
    inputs_dir = inputs_dir.resolve()
    if not inputs_dir.exists():
        print(f"Error: inputs directory not found: {inputs_dir}", file=sys.stderr)
        sys.exit(1)

    # ── Resolve output dir ───────────────────────────────────────────────────
    out_dir = Path(args.out_dir) if args.out_dir else rsd_root / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── Load config ──────────────────────────────────────────────────────────
    config_path = Path(args.config) if args.config else rsd_root / "config.json"
    if config_path.exists():
        config = RSDConfig.from_file(config_path)
    else:
        config = RSDConfig()

    # CLI overrides
    config.model_sheets = [s.strip() for s in args.sheets.split(",") if s.strip()]
    if args.max_levels is not None:
        config.max_formula_levels = args.max_levels
    if args.approximate:
        config.approximate = True

    # ── Run pipeline ─────────────────────────────────────────────────────────
    result = run(model_path, inputs_dir, config, verbose=args.verbose)

    # ── Write report ─────────────────────────────────────────────────────────
    out_path = out_dir / f"raw_sources_{model_path.stem}.xlsx"
    write_report(result, config, model_path, out_path)

    # ── Summary ──────────────────────────────────────────────────────────────
    n_found  = sum(1 for n in result.source_nodes if n.found_on_disk and n.level > 0)
    n_levels = max((r.level for r in result.formula_refs), default=0)

    print(f"\nReport: {out_path}")
    print(f"  Formula tracing : {n_levels} levels, {n_found} files found, "
          f"{len(result.missing_files)} missing")
    print(f"  Vector matching : {len(result.matched_vectors)} matched, "
          f"{len(result.unmatched_vectors)} unmatched")
    print(f"  Raw sources     : {len(result.raw_sources)} connections")


if __name__ == "__main__":
    main()
