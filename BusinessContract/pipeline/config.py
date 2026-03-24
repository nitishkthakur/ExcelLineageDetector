"""Configuration for the Business Contract pipeline."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ContractConfig:
    """All settings for generating a Business Contract."""

    model_path: Path
    output_sheets: list[str]          # sheets treated as outputs
    upstream_dir: Path | None = None
    upstream_files: list[Path] = field(default_factory=list)
    out_dir: Path = Path(".")

    # LLM settings
    llm_model: str = "claude-haiku-4-5-20251001"
    llm_batch_size: int = 20          # variables per LLM call

    # Tracing settings
    max_formula_level: int = 10
    min_vector_length: int = 3
    similarity_metric: str = "pearson"
    min_similarity: float = 0.8
    subsequence_matching: bool = True

    # Performance
    max_workers: int | None = None    # None = auto

    def __post_init__(self) -> None:
        """Validate configuration after initialization."""
        if not self.output_sheets:
            raise ValueError("output_sheets must contain at least one sheet name")
        if not isinstance(self.output_sheets, list):
            raise TypeError("output_sheets must be a list of strings")
        if not (0.0 <= self.min_similarity <= 1.0):
            raise ValueError(f"min_similarity must be 0.0-1.0, got {self.min_similarity}")
        if self.max_formula_level < 1:
            raise ValueError(f"max_formula_level must be >= 1, got {self.max_formula_level}")
        if self.min_vector_length < 1:
            raise ValueError(f"min_vector_length must be >= 1, got {self.min_vector_length}")
        if self.llm_batch_size < 1:
            raise ValueError(f"llm_batch_size must be >= 1, got {self.llm_batch_size}")

    @classmethod
    def from_file(cls, path: Path) -> ContractConfig:
        with open(path) as f:
            d = json.load(f)
        # Validate required fields
        for key in ("model_path", "output_sheets"):
            if key not in d:
                raise ValueError(f"Config file missing required key: {key}")
        d["model_path"] = Path(d["model_path"])
        if d.get("upstream_dir"):
            d["upstream_dir"] = Path(d["upstream_dir"])
        if d.get("upstream_files"):
            d["upstream_files"] = [Path(p) for p in d["upstream_files"]]
        if d.get("out_dir"):
            d["out_dir"] = Path(d["out_dir"])
        return cls(**d)
