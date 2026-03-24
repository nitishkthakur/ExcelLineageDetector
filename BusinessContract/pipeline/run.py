"""Main orchestrator for Business Contract generation."""
from __future__ import annotations

import logging
import time
from pathlib import Path

from .config import ContractConfig
from .models import BusinessContract

logger = logging.getLogger(__name__)


def generate_contract(
    config: ContractConfig,
    skip_llm: bool = False,
) -> BusinessContract:
    """Generate a complete Business Contract.

    Steps:
    1. Scan model file (vectors, formulas, connections)
    2. Enrich with upstream lineage (value-based matching)
    3. Convert formulas to SQL notation
    4. Infer business names via LLM (optional)
    5. Write Business Contract Excel
    6. Generate Mermaid diagrams
    7. Generate Python refactor code
    """
    timings: dict[str, float] = {}
    t0 = time.time()

    # 1. Scan model
    t_step = time.time()
    logger.info("Scanning model file: %s", config.model_path)
    from .scanner import scan_model
    contract = scan_model(config)
    timings["scan"] = time.time() - t_step
    logger.info(
        "Found %d variables, %d transformations, %d connections (%.1fs)",
        len(contract.variables), len(contract.transformations),
        len(contract.connections), timings["scan"],
    )

    # 2. Upstream lineage
    if config.upstream_dir or config.upstream_files:
        t_step = time.time()
        logger.info("Tracing upstream lineage...")
        try:
            from .graph_builder import enrich_upstream
            contract = enrich_upstream(contract, config)
            upstream_count = sum(1 for v in contract.variables if v.upstream_source)
            logger.info("Matched %d variables to upstream sources", upstream_count)
        except Exception as e:
            logger.warning("Upstream tracing failed (continuing without): %s", e)
        timings["upstream"] = time.time() - t_step

    # 3. Convert formulas to SQL
    t_step = time.time()
    logger.info("Converting formulas to SQL notation...")
    from .formula_converter import excel_to_sql

    var_names = {
        f"{v.sheet}!{v.cell_range}": v.business_name or v.id
        for v in contract.variables
    }
    for tx in contract.transformations:
        tx.sql_formula = excel_to_sql(tx.excel_formula, var_names)
    timings["sql_convert"] = time.time() - t_step

    # 4. LLM business names
    if not skip_llm:
        t_step = time.time()
        logger.info("Inferring business names via LLM...")
        try:
            from .llm_namer import infer_business_names
            contract.variables = infer_business_names(
                config.model_path,
                contract.variables,
                llm_model=config.llm_model,
                batch_size=config.llm_batch_size,
            )
            # Re-convert SQL with business names
            var_names = {
                f"{v.sheet}!{v.cell_range}": v.business_name
                for v in contract.variables
            }
            for tx in contract.transformations:
                tx.sql_formula = excel_to_sql(tx.excel_formula, var_names)
        except Exception as e:
            logger.warning("LLM naming failed (using fallback names): %s", e)
            for var in contract.variables:
                if not var.business_name:
                    var.business_name = (
                        f"{var.sheet}_{var.cell_range}".lower().replace(":", "_to_")
                    )
        timings["llm"] = time.time() - t_step
    else:
        for var in contract.variables:
            if not var.business_name:
                var.business_name = (
                    f"{var.sheet}_{var.cell_range}".lower().replace(":", "_to_")
                )

    # 5. Write Business Contract Excel
    t_step = time.time()
    out_dir = config.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    model_name = config.model_path.stem

    from .contract_writer import write_contract
    contract_path = out_dir / f"business_contract_{model_name}.xlsx"
    write_contract(contract, contract_path)
    logger.info("Wrote Business Contract: %s", contract_path)
    timings["write_excel"] = time.time() - t_step

    # 6. Mermaid diagrams
    t_step = time.time()
    try:
        from ..mermaid.generator import write_mermaid
        source_path, var_path = write_mermaid(contract, out_dir, model_name)
        logger.info("Wrote Mermaid diagrams: %s, %s", source_path, var_path)
    except Exception as e:
        logger.warning("Mermaid generation failed: %s", e)
    timings["mermaid"] = time.time() - t_step

    # 7. Python refactor
    t_step = time.time()
    try:
        from ..refactor.generator import generate_python
        py_path = out_dir / f"calculation_engine_{model_name}.py"
        generate_python(contract, py_path)
        logger.info("Wrote Python refactor: %s", py_path)
    except Exception as e:
        logger.warning("Python refactor generation failed: %s", e)
    timings["refactor"] = time.time() - t_step

    elapsed = time.time() - t0
    timing_str = ", ".join(f"{k}={v:.1f}s" for k, v in timings.items())
    logger.info("Done in %.1fs [%s]", elapsed, timing_str)

    return contract
