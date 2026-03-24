"""Configuration for RawSourcesDetection pipeline."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class RSDConfig:
    """All settings for raw source detection.

    Serialisation format (config.json):
    {
        "model_sheets": ["Sheet1", "Inputs"],
        "max_formula_levels": 5,
        "matching": {
            "exact": true,
            "approximate": false,
            "exact_decimal_places": 8,
            "subsequence_matching": true,
            "min_similarity": 0.85,
            "similarity_metric": "pearson"
        },
        "performance": {
            "max_workers": null,
            "min_vector_length": 3
        }
    }
    """

    # Sheets to trace in the model file (formula tracing + vector matching)
    model_sheets: list[str] = field(default_factory=list)

    # Formula tracing depth cap
    max_formula_levels: int = 5

    # Matching
    exact: bool = True
    approximate: bool = False          # default: exact only for speed
    exact_decimal_places: int = 8      # floating-point rounding for exact match
    subsequence_matching: bool = True  # allow model vector as sub-sequence of upstream
    min_similarity: float = 0.85       # threshold for approximate matches
    similarity_metric: str = "pearson" # pearson | cosine | euclidean

    # Performance
    max_workers: int | None = None     # None = os.cpu_count()
    min_vector_length: int = 3

    def __post_init__(self) -> None:
        if self.max_formula_levels < 1:
            raise ValueError(
                f"max_formula_levels must be >= 1, got {self.max_formula_levels}"
            )
        if not (0.0 <= self.min_similarity <= 1.0):
            raise ValueError(
                f"min_similarity must be 0.0-1.0, got {self.min_similarity}"
            )
        if self.exact_decimal_places < 0:
            raise ValueError(
                f"exact_decimal_places must be >= 0, got {self.exact_decimal_places}"
            )
        if self.min_vector_length < 1:
            raise ValueError(
                f"min_vector_length must be >= 1, got {self.min_vector_length}"
            )

    @classmethod
    def from_file(cls, path: Path) -> "RSDConfig":
        """Load config from a JSON file."""
        data = json.loads(path.read_text())
        m = data.get("matching", {})
        p = data.get("performance", {})
        return cls(
            model_sheets=data.get("model_sheets", []),
            max_formula_levels=data.get("max_formula_levels", 5),
            exact=m.get("exact", True),
            approximate=m.get("approximate", False),
            exact_decimal_places=m.get("exact_decimal_places", 8),
            subsequence_matching=m.get("subsequence_matching", True),
            min_similarity=m.get("min_similarity", 0.85),
            similarity_metric=m.get("similarity_metric", "pearson"),
            max_workers=p.get("max_workers", None),
            min_vector_length=p.get("min_vector_length", 3),
        )

    def to_trace_config(self):
        """Convert to lineage.tracing.config.TraceConfig for reuse."""
        from lineage.tracing.config import TraceConfig
        return TraceConfig(
            exact=self.exact,
            approximate=self.approximate,
            exact_decimal_places=self.exact_decimal_places,
            subsequence_matching=self.subsequence_matching,
            min_similarity=self.min_similarity,
            similarity_metric=self.similarity_metric,
            max_workers=self.max_workers,
            min_vector_length=self.min_vector_length,
        )
