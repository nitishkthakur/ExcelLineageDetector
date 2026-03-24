"""Generate Python code that reproduces the Excel calculation engine."""
from __future__ import annotations

from pathlib import Path
from textwrap import dedent, indent

from ..pipeline.models import BusinessContract, ContractVariable, TransformationStep


def generate_python(contract: BusinessContract, out_path: Path) -> Path:
    """Generate a standalone Python module from the Business Contract.

    The generated code:
    - Defines input variables as numpy arrays / scalars
    - Implements each transformation as a function
    - Computes outputs by chaining transformations
    """
    var_map = {v.id: v for v in contract.variables}

    inputs = [v for v in contract.variables if v.variable_type == "input"]
    outputs = [v for v in contract.variables if v.variable_type == "output"]

    # Topological sort of transformations
    sorted_tx = _topo_sort(contract.transformations, var_map)

    lines = [
        '"""Auto-generated calculation engine from Business Contract.',
        f'Source model: {contract.model_file}',
        '"""',
        "from __future__ import annotations",
        "",
        "import numpy as np",
        "",
        "",
        "# ── Input Variables ──────────────────────────────────────────",
        "",
    ]

    # Input declarations
    for var in inputs:
        name = _var_name(var)
        if var.length > 1:
            vals = ", ".join(str(v) for v in var.sample_values[:5])
            lines.append(f"# {var.excel_location} ({var.direction}, length={var.length})")
            lines.append(f"# Sample: [{vals}, ...]")
            lines.append(f"{name} = np.zeros({var.length})  # TODO: load actual data")
        else:
            val = var.sample_values[0] if var.sample_values else 0.0
            lines.append(f"# {var.excel_location}")
            lines.append(f"{name} = {val}")
        lines.append("")

    lines.append("")
    lines.append("# ── Transformations ─────────────────────────────────────────")
    lines.append("")

    # Transformation functions
    for tx in sorted_tx:
        out_var = var_map.get(tx.output_variable_id)
        if not out_var:
            continue

        func_name = f"compute_{_var_name(out_var)}"
        input_vars = [var_map[iid] for iid in tx.input_variable_ids if iid in var_map]
        param_names = [_var_name(v) for v in input_vars]

        lines.append(f"def {func_name}({', '.join(param_names)}):")
        lines.append(f'    """')
        lines.append(f"    Excel: {tx.excel_formula}")
        if tx.sql_formula:
            lines.append(f"    SQL:   {tx.sql_formula}")
        lines.append(f"    Sheet: {tx.sheet}, Cell: {tx.cell_range}")
        lines.append(f'    """')
        lines.append(f"    # TODO: implement transformation logic")
        lines.append(f"    # Excel formula: {tx.excel_formula}")
        if param_names:
            lines.append(f"    return {param_names[0]}  # placeholder")
        else:
            lines.append(f"    return 0.0  # placeholder")
        lines.append("")
        lines.append("")

    # Main computation chain
    lines.append("# ── Compute Outputs ─────────────────────────────────────────")
    lines.append("")
    lines.append("def compute_all():")
    lines.append('    """Run the full calculation chain."""')

    for tx in sorted_tx:
        out_var = var_map.get(tx.output_variable_id)
        if not out_var:
            continue
        func_name = f"compute_{_var_name(out_var)}"
        input_vars = [var_map[iid] for iid in tx.input_variable_ids if iid in var_map]
        param_names = [_var_name(v) for v in input_vars]
        out_name = _var_name(out_var)
        lines.append(f"    {out_name} = {func_name}({', '.join(param_names)})")

    lines.append("")
    lines.append("    # Return outputs")
    lines.append("    return {")
    for var in outputs:
        name = _var_name(var)
        lines.append(f'        "{name}": {name},')
    lines.append("    }")
    lines.append("")
    lines.append("")
    lines.append('if __name__ == "__main__":')
    lines.append("    results = compute_all()")
    lines.append('    for name, value in results.items():')
    lines.append('        print(f"{name}: {value}")')
    lines.append("")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines))
    return out_path


def _var_name(var: ContractVariable) -> str:
    """Get a valid Python variable name."""
    name = var.business_name or f"{var.sheet}_{var.cell_range}"
    name = name.lower().replace(" ", "_").replace("-", "_").replace(":", "_")
    name = "".join(c for c in name if c.isalnum() or c == "_")
    if not name:
        name = f"v_{var.id}"
    elif name[0].isdigit():
        name = f"v_{name}"
    return name


def _topo_sort(
    transformations: list[TransformationStep],
    var_map: dict[str, ContractVariable],
) -> list[TransformationStep]:
    """Topological sort with cycle detection.

    Uses a three-state DFS (unvisited / in-progress / done) to detect cycles.
    Cycles are broken by dropping the back-edge dependency.
    """
    tx_by_output: dict[str, TransformationStep] = {}
    for tx in transformations:
        tx_by_output[tx.output_variable_id] = tx

    UNVISITED, IN_PROGRESS, DONE = 0, 1, 2
    state: dict[str, int] = {tx.output_variable_id: UNVISITED for tx in transformations}
    result: list[TransformationStep] = []

    def visit(tx: TransformationStep) -> None:
        vid = tx.output_variable_id
        if state.get(vid) == DONE:
            return
        if state.get(vid) == IN_PROGRESS:
            # Cycle detected — skip to break the cycle
            return
        state[vid] = IN_PROGRESS
        for input_id in tx.input_variable_ids:
            if input_id in tx_by_output:
                visit(tx_by_output[input_id])
        state[vid] = DONE
        result.append(tx)

    for tx in transformations:
        if state.get(tx.output_variable_id) == UNVISITED:
            visit(tx)

    return result
