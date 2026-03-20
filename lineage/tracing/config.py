"""Configuration for upstream tracing."""
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass
class TraceConfig:
    """Configuration for upstream vector tracing."""

    # Matching modes
    exact: bool = True
    approximate: bool = True
    top_n: int = 5          # top-N approximate matches (ignored if approximate=False)

    # Exact matching
    exact_decimal_places: int = 8
    subsequence_matching: bool = True

    # Approximate matching
    similarity_metric: str = "pearson"   # pearson | cosine | euclidean
    min_similarity: float = 0.8
    length_tolerance_pct: float = 50.0   # allow +/- this % length mismatch
    direction_sensitive: bool = False

    # Performance
    max_workers: int | None = None       # None = os.cpu_count()
    min_vector_length: int = 3

    # ------------------------------------------------------------------ #
    # Serialization
    # ------------------------------------------------------------------ #

    @classmethod
    def from_file(cls, path: Path) -> TraceConfig:
        """Load config from a JSON or YAML file."""
        text = path.read_text()
        suffix = path.suffix.lower()

        if suffix in (".yaml", ".yml"):
            try:
                import yaml
                data = yaml.safe_load(text) or {}
            except ImportError:
                raise ImportError("pyyaml required for YAML config: pip install pyyaml")
        else:
            data = json.loads(text)

        m = data.get("matching", {})
        p = data.get("performance", {})

        return cls(
            exact=m.get("exact", cls.exact),
            approximate=m.get("approximate", cls.approximate),
            top_n=m.get("top_n", cls.top_n),
            exact_decimal_places=m.get("exact_decimal_places", cls.exact_decimal_places),
            subsequence_matching=m.get("subsequence_matching", cls.subsequence_matching),
            similarity_metric=m.get("similarity_metric", cls.similarity_metric),
            min_similarity=m.get("min_similarity", cls.min_similarity),
            length_tolerance_pct=m.get("length_tolerance_pct", cls.length_tolerance_pct),
            direction_sensitive=m.get("direction_sensitive", cls.direction_sensitive),
            max_workers=p.get("max_workers", cls.max_workers),
            min_vector_length=p.get("min_vector_length", cls.min_vector_length),
        )

    def to_dict(self) -> dict:
        return {
            "matching": {
                "exact": self.exact,
                "approximate": self.approximate,
                "top_n": self.top_n,
                "exact_decimal_places": self.exact_decimal_places,
                "subsequence_matching": self.subsequence_matching,
                "similarity_metric": self.similarity_metric,
                "min_similarity": self.min_similarity,
                "length_tolerance_pct": self.length_tolerance_pct,
                "direction_sensitive": self.direction_sensitive,
            },
            "performance": {
                "max_workers": self.max_workers,
                "min_vector_length": self.min_vector_length,
            },
        }
