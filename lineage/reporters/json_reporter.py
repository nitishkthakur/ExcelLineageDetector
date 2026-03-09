"""JSON reporter for Excel Lineage Detector."""

from __future__ import annotations
import json
from collections import Counter
from datetime import datetime
from pathlib import Path

from lineage.models import DataConnection


class JsonReporter:
    """Writes lineage results as a JSON file."""

    def write(
        self,
        connections: list[DataConnection],
        input_path: Path,
        out_dir: Path,
    ) -> Path:
        """Write connections to a JSON file.

        Args:
            connections: List of detected DataConnection objects.
            input_path: Path to the analyzed Excel file.
            out_dir: Directory to write the output file.

        Returns:
            Path to the written JSON file.
        """
        stem = input_path.stem
        out = out_dir / f"{stem}_lineage.json"

        by_category = dict(Counter(c.category for c in connections))
        by_subtype = dict(Counter(c.sub_type for c in connections))

        data = {
            "file": str(input_path),
            "scanned_at": datetime.utcnow().isoformat() + "Z",
            "summary": {
                "total_connections": len(connections),
                "by_category": by_category,
                "by_subtype": by_subtype,
            },
            "connections": [c.to_dict() for c in connections],
        }

        out.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")
        return out
