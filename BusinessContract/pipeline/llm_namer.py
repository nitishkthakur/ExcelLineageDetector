"""LLM-based business name inference using LangChain + MCP tools."""
from __future__ import annotations

import json
import logging
import zipfile
from pathlib import Path

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage, SystemMessage

from ..mcp_server.streaming import (
    _load_shared_strings,
    get_sheet_summary,
    read_cell_neighborhood,
)
from .models import ContractVariable

logger = logging.getLogger(__name__)

_SYSTEM_PROMPT = """You are an expert financial analyst. Your task is to assign short,
precise business names to data vectors found in an Excel workbook.

For each variable, you will be given:
- Its sheet name and cell range
- The cells around it (headers, labels, nearby values)
- Its direction (row or column vector)
- Sample values

Rules:
1. Use concise business terminology (e.g., "quarterly_revenue", "discount_rate", "unit_price")
2. Use snake_case for names
3. Names should be self-documenting — someone reading just the name should understand what the data is
4. If the variable appears to be a time series, include the frequency (e.g., "monthly_sales", "daily_prices")
5. If values look like percentages (0-1 range), reflect that (e.g., "growth_rate", "tax_rate")

Respond with a JSON array of objects: [{"id": "<variable_id>", "business_name": "<name>"}]
Only output valid JSON, no markdown fences or extra text."""


def _gather_context_for_variable(
    zf: zipfile.ZipFile,
    shared_strings: list[str],
    var: ContractVariable,
    radius: int = 3,
) -> str:
    """Gather neighborhood context for a variable to send to LLM."""
    # Get the start cell of the range
    start_cell = var.cell_range.split(":")[0]

    cells = read_cell_neighborhood(
        zf, var.sheet, start_cell, radius=radius,
        shared_strings=shared_strings,
    )

    # Format as readable text
    lines = [
        f"Sheet: {var.sheet}",
        f"Range: {var.cell_range} (direction: {var.direction}, length: {var.length})",
        f"Sample values: {var.sample_values}",
        f"Source type: {var.source_type}",
        "",
        "Nearby cells:",
    ]
    for ref, cell in sorted(cells.items()):
        val_str = cell.value or ""
        if cell.formula:
            val_str += f" [formula: {cell.formula}]"
        lines.append(f"  {ref}: {val_str}")

    return "\n".join(lines)


def infer_business_names(
    model_path: Path,
    variables: list[ContractVariable],
    llm_model: str = "claude-haiku-4-5-20251001",
    batch_size: int = 20,
) -> list[ContractVariable]:
    """Use LLM to infer business names for variables.

    Reads cell neighborhoods directly (no MCP server needed for batch mode).
    Updates variables in-place and returns them.
    """
    if not variables:
        return variables

    llm = ChatAnthropic(model=llm_model, temperature=0, max_tokens=4096)

    with zipfile.ZipFile(str(model_path), "r") as zf:
        shared_strings = _load_shared_strings(zf)

        # Process in batches
        for i in range(0, len(variables), batch_size):
            batch = variables[i : i + batch_size]

            # Gather context for each variable
            contexts = []
            for var in batch:
                ctx = _gather_context_for_variable(zf, shared_strings, var)
                contexts.append(f"Variable {var.id}:\n{ctx}")

            prompt = (
                "Assign business names to the following variables:\n\n"
                + "\n\n---\n\n".join(contexts)
            )

            messages = [
                SystemMessage(content=_SYSTEM_PROMPT),
                HumanMessage(content=prompt),
            ]

            batch_num = i // batch_size + 1
            max_retries = 2
            success = False

            for attempt in range(max_retries + 1):
                try:
                    response = llm.invoke(messages)
                    content = response.content

                    # Parse JSON response — handle markdown fences
                    if "```" in content:
                        # Extract content between first pair of fences
                        parts = content.split("```")
                        if len(parts) >= 3:
                            inner = parts[1]
                            if inner.startswith("json"):
                                inner = inner[4:]
                            content = inner

                    names = json.loads(content.strip())
                    if not isinstance(names, list):
                        raise ValueError(f"Expected JSON array, got {type(names)}")

                    name_map = {n["id"]: n["business_name"] for n in names
                                if isinstance(n, dict) and "id" in n and "business_name" in n}

                    for var in batch:
                        if var.id in name_map:
                            var.business_name = name_map[var.id]

                    logger.info(
                        "Named %d/%d variables in batch %d",
                        len(name_map), len(batch), batch_num,
                    )
                    success = True
                    break

                except Exception as e:
                    if attempt < max_retries:
                        logger.warning(
                            "LLM naming batch %d attempt %d failed: %s — retrying",
                            batch_num, attempt + 1, e,
                        )
                    else:
                        logger.warning(
                            "LLM naming failed for batch %d after %d attempts: %s",
                            batch_num, max_retries + 1, e,
                        )

            if not success:
                # Fallback: context-aware names from headers/sheet
                for var in batch:
                    if not var.business_name:
                        var.business_name = _fallback_name(var)

    # Ensure all variables have names
    for var in variables:
        if not var.business_name:
            var.business_name = _fallback_name(var)

    return variables


def _fallback_name(var: ContractVariable) -> str:
    """Generate a descriptive fallback name when LLM is unavailable."""
    sheet = var.sheet.lower().replace(" ", "_")
    rng = var.cell_range.replace(":", "_to_").replace("$", "")

    # Use direction for context
    if var.direction == "scalar":
        return f"{sheet}_{rng}_scalar"
    elif var.direction == "column":
        return f"{sheet}_{rng}_col_vec"
    elif var.direction == "row":
        return f"{sheet}_{rng}_row_vec"
    return f"{sheet}_{rng}"
