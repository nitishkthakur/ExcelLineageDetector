"""Exact matching engine for upstream tracing.

Supports:
1. Full-vector exact match  — hash-based O(1) lookup
2. Subsequence exact match  — batched numpy sliding-window comparison,
   grouped by upstream length for vectorized processing.
"""
from __future__ import annotations

import numpy as np

from lineage.tracing.config import TraceConfig
from lineage.tracing.models import TracingVector, VectorMatch
from lineage.tracing.scanner import compute_sub_range

# Memory guard for batch subsequence matching.
_MAX_BATCH_ELEMENTS = 10_000_000


class ExactMatcher:
    """Hash-based exact matching with batched subsequence support."""

    def __init__(self, config: TraceConfig):
        self.config = config
        self.dp = config.exact_decimal_places
        # rounded-value-tuple  →  list of upstream TracingVectors
        self._index: dict[tuple[float, ...], list[TracingVector]] = {}
        # For subsequence: upstream grouped by length, pre-rounded numpy arrays
        self._by_length: dict[int, tuple[list[TracingVector], np.ndarray]] = {}

    # ------------------------------------------------------------------ #
    # Indexing
    # ------------------------------------------------------------------ #

    def index_upstream(self, vectors: list[TracingVector]) -> None:
        """Add upstream vectors to the index."""
        # Hash index for full-vector match
        for vec in vectors:
            key = tuple(round(v, self.dp) for v in vec.values)
            self._index.setdefault(key, []).append(vec)

        # Group by length and pre-stack rounded arrays for subsequence matching
        if self.config.subsequence_matching:
            by_len: dict[int, list[TracingVector]] = {}
            for v in vectors:
                by_len.setdefault(v.length, []).append(v)
            self._by_length = {
                length: (
                    vecs,
                    np.round(np.array([v.values for v in vecs], dtype=np.float64), self.dp),
                )
                for length, vecs in by_len.items()
            }

    # ------------------------------------------------------------------ #
    # Matching
    # ------------------------------------------------------------------ #

    def match(self, model_vec: TracingVector) -> list[VectorMatch]:
        """Return all exact matches for *model_vec*."""
        matches: list[VectorMatch] = []
        seen: set[str] = set()

        model_key = tuple(round(v, self.dp) for v in model_vec.values)

        # --- 1. Full-vector exact match (hash lookup, O(1)) ---
        for upstream in self._index.get(model_key, []):
            uid = f"{upstream.file}|{upstream.sheet}|{upstream.cell_range}"
            if uid in seen:
                continue
            seen.add(uid)
            matches.append(self._make_match(
                model_vec, upstream, "exact", 1.0, upstream.cell_range,
            ))

        # --- 2. Subsequence match (batched per-length) ---
        if self.config.subsequence_matching:
            model_len = model_vec.length
            model_rounded = np.round(
                np.array(model_vec.values, dtype=np.float64), self.dp,
            )

            for up_len, (vecs, batch) in self._by_length.items():
                if up_len <= model_len:
                    continue

                if self.config.direction_sensitive:
                    keep = [v.direction == model_vec.direction for v in vecs]
                    if not any(keep):
                        continue
                    keep_arr = np.array(keep)
                    vecs_f = [v for v, k in zip(vecs, keep) if k]
                    batch_f = batch[keep_arr]
                else:
                    vecs_f = vecs
                    batch_f = batch

                n_vecs = len(vecs_f)
                n_win = up_len - model_len + 1
                total_els = n_vecs * n_win * model_len

                if total_els <= _MAX_BATCH_ELEMENTS:
                    # Batch: 2D sliding window view, then flat comparison
                    # batch_f shape: (N, up_len)
                    windows_3d = np.lib.stride_tricks.sliding_window_view(
                        batch_f, model_len, axis=1,
                    )  # (N, n_win, model_len)

                    # Compare each window against model_rounded: broadcast
                    # (N, n_win, model_len) == (model_len,) -> (N, n_win, model_len)
                    eq = windows_3d == model_rounded
                    match_2d = np.all(eq, axis=2)  # (N, n_win)

                    # Find hits
                    vec_idxs, win_idxs = np.where(match_2d)
                    for vi, wi in zip(vec_idxs, win_idxs):
                        vi, wi = int(vi), int(wi)
                        v = vecs_f[vi]
                        sub_range = compute_sub_range(v, wi, model_len)
                        uid = f"{v.file}|{v.sheet}|{sub_range}"
                        if uid in seen:
                            continue
                        seen.add(uid)
                        matches.append(self._make_match(
                            model_vec, v, "exact_subsequence", 1.0, sub_range,
                        ))
                else:
                    # Per-vector fallback
                    for vi, v in enumerate(vecs_f):
                        up_rounded = batch_f[vi]
                        windows = np.lib.stride_tricks.sliding_window_view(
                            up_rounded, model_len,
                        )
                        eq_mask = np.all(windows == model_rounded, axis=1)
                        for offset in np.where(eq_mask)[0]:
                            sub_range = compute_sub_range(v, int(offset), model_len)
                            uid = f"{v.file}|{v.sheet}|{sub_range}"
                            if uid in seen:
                                continue
                            seen.add(uid)
                            matches.append(self._make_match(
                                model_vec, v, "exact_subsequence", 1.0, sub_range,
                            ))

        return matches

    # ------------------------------------------------------------------ #
    # Helpers
    # ------------------------------------------------------------------ #

    @staticmethod
    def _make_match(
        model: TracingVector,
        upstream: TracingVector,
        match_type: str,
        similarity: float,
        matched_range: str,
    ) -> VectorMatch:
        return VectorMatch(
            model_sheet=model.sheet,
            model_range=model.cell_range,
            model_direction=model.direction,
            model_length=model.length,
            model_sample=list(model.values[:5]),
            match_rank=0,
            match_type=match_type,
            similarity=similarity,
            upstream_file=upstream.file,
            upstream_sheet=upstream.sheet,
            upstream_range=upstream.cell_range,
            upstream_direction=upstream.direction,
            upstream_length=upstream.length,
            upstream_sample=list(upstream.values[:5]),
            upstream_matched_range=matched_range,
        )
