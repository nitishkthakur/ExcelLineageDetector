"""Build dependency graph and enrich with upstream lineage."""
from __future__ import annotations

import logging
from pathlib import Path

from lineage.tracing.config import TraceConfig
from lineage.tracing.tracer import UpstreamTracer

from .config import ContractConfig
from .models import BusinessContract, ContractVariable, DependencyEdge

logger = logging.getLogger(__name__)


def enrich_upstream(
    contract: BusinessContract,
    config: ContractConfig,
) -> BusinessContract:
    """Trace hardcoded input vectors back to upstream files.

    Uses the existing UpstreamTracer for value-based matching.
    Updates variables with upstream_source, upstream_file, etc.
    Adds vector_match edges.
    """
    if not config.upstream_dir and not config.upstream_files:
        return contract

    # Collect upstream file paths (exclude the model file itself)
    model_resolved = config.model_path.resolve()
    upstream_paths: list[Path] = list(config.upstream_files)
    if config.upstream_dir and config.upstream_dir.is_dir():
        for pattern in ("*.xlsx", "*.xlsm"):
            for p in config.upstream_dir.rglob(pattern):
                if p.resolve() != model_resolved and p not in upstream_paths:
                    upstream_paths.append(p)

    if not upstream_paths:
        return contract

    trace_cfg = TraceConfig(
        similarity_metric=config.similarity_metric,
        min_similarity=config.min_similarity,
        subsequence_matching=config.subsequence_matching,
        min_vector_length=config.min_vector_length,
    )

    tracer = UpstreamTracer(trace_cfg)

    # Run tracing for each sheet that has hardcoded inputs
    input_vars = [v for v in contract.variables if v.source_type == "hardcoded"]
    sheets_with_inputs = {v.sheet for v in input_vars}

    for sheet in sheets_with_inputs:
        try:
            matches, _unmatched = tracer.trace(
                model_path=config.model_path,
                sheet_name=sheet,
                upstream_paths=upstream_paths,
            )
        except Exception as e:
            logger.warning("Upstream tracing failed for sheet %s: %s", sheet, e)
            continue

        # Match results back to our variables
        for match in matches:
            for var in input_vars:
                if var.sheet == sheet and _ranges_overlap(
                    var.cell_range, match.model_range
                ):
                    var.upstream_source = (
                        f"{match.upstream_file}!{match.upstream_sheet}!"
                        f"{match.upstream_range}"
                    )
                    var.upstream_file = str(match.upstream_file)
                    var.upstream_sheet = match.upstream_sheet
                    var.upstream_range = match.upstream_range
                    var.confidence = match.similarity
                    var.match_type = match.match_type
                    var.source_type = "external_link"

                    contract.edges.append(DependencyEdge(
                        source_id=f"upstream_{match.upstream_file}_{match.upstream_sheet}",
                        target_id=var.id,
                        edge_type="vector_match",
                        metadata={
                            "upstream_file": str(match.upstream_file),
                            "upstream_sheet": match.upstream_sheet,
                            "upstream_range": match.upstream_range,
                            "similarity": match.similarity,
                            "match_type": match.match_type,
                        },
                    ))

    return contract


def _ranges_overlap(range1: str, range2: str) -> bool:
    """Check if two cell ranges overlap using proper rectangle intersection."""
    from .scanner import _parse_range
    r1 = _parse_range(range1)
    r2 = _parse_range(range2)
    if not r1 or not r2:
        # Fallback to normalized string comparison
        s1 = range1.replace("$", "")
        s2 = range2.replace("$", "")
        return s1 == s2 or s1 in s2 or s2 in s1
    return not (r1[2] < r2[0] or r1[0] > r2[2] or r1[3] < r2[1] or r1[1] > r2[3])
