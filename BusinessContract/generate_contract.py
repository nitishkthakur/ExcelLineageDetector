#!/usr/bin/env python3
"""CLI: Generate a Business Contract from an Excel model file."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

# Ensure project root is on path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dotenv import load_dotenv

from BusinessContract.pipeline.config import ContractConfig
from BusinessContract.pipeline.run import generate_contract

load_dotenv(Path(__file__).parent / ".env")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a Business Contract from an Excel model file"
    )
    parser.add_argument("model", type=Path, help="Path to the model Excel file")
    parser.add_argument(
        "--output-sheets", "-o", nargs="+", required=True,
        help="Sheet name(s) treated as outputs",
    )
    parser.add_argument(
        "--upstream-dir", "-u", type=Path, default=None,
        help="Directory containing upstream Excel files",
    )
    parser.add_argument(
        "--out-dir", "-d", type=Path, default=Path("./contract_output"),
        help="Output directory (default: ./contract_output)",
    )
    parser.add_argument(
        "--skip-llm", action="store_true",
        help="Skip LLM business name inference (use fallback names)",
    )
    parser.add_argument(
        "--llm-model", default="claude-haiku-4-5-20251001",
        help="LLM model for business name inference",
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="Enable verbose logging",
    )

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    config = ContractConfig(
        model_path=args.model.resolve(),
        output_sheets=args.output_sheets,
        upstream_dir=args.upstream_dir.resolve() if args.upstream_dir else None,
        out_dir=args.out_dir.resolve(),
        llm_model=args.llm_model,
    )

    contract = generate_contract(config, skip_llm=args.skip_llm)

    print(f"\nBusiness Contract generated:")
    print(f"  Variables:       {len(contract.variables)}")
    print(f"  Transformations: {len(contract.transformations)}")
    print(f"  Connections:     {len(contract.connections)}")
    print(f"  Output dir:      {config.out_dir}")


if __name__ == "__main__":
    main()
