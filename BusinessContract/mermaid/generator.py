"""Generate Mermaid flowchart diagrams from a Business Contract."""
from __future__ import annotations

import re
from pathlib import Path

from ..pipeline.models import BusinessContract, ContractVariable


def _sanitize(text: str) -> str:
    """Sanitize text for Mermaid node labels."""
    return re.sub(r'["\[\]{}|<>]', "", text).replace("\n", " ").strip()[:60]


def _node_id(var: ContractVariable) -> str:
    """Create a valid Mermaid node ID."""
    return f"v_{var.id}"


def generate_source_level(contract: BusinessContract) -> str:
    """Generate a source-level Mermaid diagram.

    Shows files/sources as nodes and data flow between them.
    Upstream files -> Model sheets -> Output sheets.
    Uses O(1) dict lookups instead of list.index().
    """
    lines = ["graph LR"]

    # Collect unique sources
    upstream_files: set[str] = set()
    model_sheets: set[str] = set()
    output_sheets: set[str] = set(contract.output_sheets)

    for var in contract.variables:
        model_sheets.add(var.sheet)
        if var.upstream_file:
            upstream_files.add(var.upstream_file)

    for conn in contract.connections:
        raw = conn.get("raw_connection", "")
        if raw:
            upstream_files.add(raw[:60])

    uf_list = sorted(upstream_files)
    ms_list = sorted(model_sheets - output_sheets)
    os_list = sorted(output_sheets)

    # Build index maps for O(1) lookup
    uf_idx = {uf: i for i, uf in enumerate(uf_list)}
    ms_idx = {ms: i for i, ms in enumerate(ms_list)}
    os_idx = {os: i for i, os in enumerate(os_list)}

    # Nodes
    for i, uf in enumerate(uf_list):
        name = Path(uf).name if "/" in uf or "\\" in uf else uf
        safe = _sanitize(name)
        lines.append(f'    uf{i}[("{safe}")]')
    for i, sheet in enumerate(ms_list):
        lines.append(f'    ms{i}["{_sanitize(sheet)}"]')
    for i, sheet in enumerate(os_list):
        lines.append(f'    os{i}[["{_sanitize(sheet)}"]]')

    # Edges (deduplicated via set)
    seen_edges: set[str] = set()

    for var in contract.variables:
        if var.upstream_file and var.upstream_file in uf_idx:
            src = f"uf{uf_idx[var.upstream_file]}"
            if var.sheet in ms_idx:
                tgt = f"ms{ms_idx[var.sheet]}"
            elif var.sheet in os_idx:
                tgt = f"os{os_idx[var.sheet]}"
            else:
                continue
            edge_key = f"{src}->{tgt}"
            if edge_key not in seen_edges:
                seen_edges.add(edge_key)
                lines.append(f"    {src} --> {tgt}")

    # Build var_id -> var map for edge resolution
    var_map = {v.id: v for v in contract.variables}

    for edge in contract.edges:
        if edge.edge_type != "formula":
            continue
        sv = var_map.get(edge.source_id)
        tv = var_map.get(edge.target_id)
        if not sv or not tv:
            continue

        src_node = ms_idx.get(sv.sheet)
        if src_node is None:
            continue
        src = f"ms{src_node}"

        if tv.sheet in os_idx:
            tgt = f"os{os_idx[tv.sheet]}"
        elif tv.sheet in ms_idx and ms_idx[tv.sheet] != src_node:
            tgt = f"ms{ms_idx[tv.sheet]}"
        else:
            continue

        edge_key = f"{src}->{tgt}"
        if edge_key not in seen_edges:
            seen_edges.add(edge_key)
            lines.append(f"    {src} --> {tgt}")

    # Styling
    lines.append("")
    lines.append("    classDef upstream fill:#e8f5e9,stroke:#4caf50")
    lines.append("    classDef model fill:#e3f2fd,stroke:#2196f3")
    lines.append("    classDef output fill:#fce4ec,stroke:#e91e63")

    for i in range(len(uf_list)):
        lines.append(f"    class uf{i} upstream")
    for i in range(len(ms_list)):
        lines.append(f"    class ms{i} model")
    for i in range(len(os_list)):
        lines.append(f"    class os{i} output")

    return "\n".join(lines)


def generate_variable_level(contract: BusinessContract) -> str:
    """Generate a variable-level Mermaid diagram.

    Shows individual variables as nodes with formula edges between them.
    """
    lines = ["graph TD"]

    # Group variables by type for subgraphs
    inputs = [v for v in contract.variables if v.variable_type == "input"]
    intermediates = [v for v in contract.variables if v.variable_type == "intermediate"]
    outputs = [v for v in contract.variables if v.variable_type == "output"]

    # Input subgraph
    if inputs:
        lines.append("    subgraph Inputs")
        for var in inputs:
            label = _sanitize(var.business_name or var.excel_location)
            lines.append(f'        {_node_id(var)}["{label}"]')
        lines.append("    end")

    # Intermediate subgraph
    if intermediates:
        lines.append("    subgraph Calculations")
        for var in intermediates:
            label = _sanitize(var.business_name or var.excel_location)
            lines.append(f'        {_node_id(var)}["{label}"]')
        lines.append("    end")

    # Output subgraph
    if outputs:
        lines.append("    subgraph Outputs")
        for var in outputs:
            label = _sanitize(var.business_name or var.excel_location)
            lines.append(f'        {_node_id(var)}[["{label}"]]')
        lines.append("    end")

    # Edges
    var_map = {v.id: v for v in contract.variables}
    seen_edges: set[str] = set()

    for edge in contract.edges:
        if edge.source_id in var_map and edge.target_id in var_map:
            key = f"{edge.source_id}->{edge.target_id}"
            if key not in seen_edges:
                seen_edges.add(key)
                src = _node_id(var_map[edge.source_id])
                tgt = _node_id(var_map[edge.target_id])
                label = edge.edge_type
                lines.append(f"    {src} -->|{label}| {tgt}")

    # Styling
    lines.append("")
    lines.append("    classDef input fill:#e8f5e9,stroke:#4caf50")
    lines.append("    classDef intermediate fill:#e3f2fd,stroke:#2196f3")
    lines.append("    classDef output fill:#fce4ec,stroke:#e91e63")

    for var in inputs:
        lines.append(f"    class {_node_id(var)} input")
    for var in intermediates:
        lines.append(f"    class {_node_id(var)} intermediate")
    for var in outputs:
        lines.append(f"    class {_node_id(var)} output")

    return "\n".join(lines)


def write_mermaid(
    contract: BusinessContract,
    out_dir: Path,
    model_name: str = "model",
) -> tuple[Path, Path]:
    """Write both Mermaid diagrams to markdown files.

    Returns (source_level_path, variable_level_path).
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    source_md = out_dir / f"{model_name}_source_flow.md"
    variable_md = out_dir / f"{model_name}_variable_flow.md"

    source_diagram = generate_source_level(contract)
    variable_diagram = generate_variable_level(contract)

    source_md.write_text(
        f"# Source-Level Data Flow\n\n```mermaid\n{source_diagram}\n```\n"
    )
    variable_md.write_text(
        f"# Variable-Level Data Flow\n\n```mermaid\n{variable_diagram}\n```\n"
    )

    return source_md, variable_md
