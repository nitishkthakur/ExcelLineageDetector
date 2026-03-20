"""Approximate matching engine for upstream tracing.

Supports: Pearson correlation, cosine similarity, Euclidean distance.
Uses numpy for fully vectorized batch computation — all upstream vectors of the
same length are processed in a single matrix operation.
"""
from __future__ import annotations

import numpy as np

from lineage.tracing.config import TraceConfig
from lineage.tracing.models import TracingVector, VectorMatch
from lineage.tracing.scanner import compute_sub_range

# Memory guard: max float elements in a single batch matrix.
_MAX_BATCH_ELEMENTS = 10_000_000


# ---------------------------------------------------------------------------
# Batch similarity kernels
# ---------------------------------------------------------------------------

def _batch_pearson(model: np.ndarray, batch: np.ndarray) -> np.ndarray:
    """Vectorized Pearson correlation: model (L,) vs batch (N, L) -> (N,)."""
    if batch.ndim == 1:
        batch = batch.reshape(1, -1)
    if batch.shape[1] < 2:
        return np.zeros(batch.shape[0])

    model_c = model - model.mean()
    batch_c = batch - batch.mean(axis=1, keepdims=True)

    model_ss = np.dot(model_c, model_c)
    batch_ss = np.sum(batch_c ** 2, axis=1)

    if model_ss < 1e-30:
        return np.where(
            batch_ss < 1e-30,
            np.where(np.abs(batch[:, 0] - model[0]) < 1e-10, 1.0, 0.0),
            0.0,
        )

    numerator = batch_c @ model_c
    denominator = np.sqrt(batch_ss * model_ss)
    safe = denominator > 1e-15
    result = np.where(safe, numerator / np.where(safe, denominator, 1.0), 0.0)
    return np.clip(result, -1.0, 1.0)


def _batch_cosine(model: np.ndarray, batch: np.ndarray) -> np.ndarray:
    """Vectorized cosine similarity: model (L,) vs batch (N, L) -> (N,)."""
    if batch.ndim == 1:
        batch = batch.reshape(1, -1)

    model_norm = np.linalg.norm(model)
    if model_norm < 1e-15:
        return np.zeros(batch.shape[0])

    batch_norms = np.linalg.norm(batch, axis=1)
    dots = batch @ model
    denoms = batch_norms * model_norm
    safe = denoms > 1e-15
    result = np.where(safe, dots / np.where(safe, denoms, 1.0), 0.0)
    return np.clip(result, -1.0, 1.0)


def _batch_euclidean(model: np.ndarray, batch: np.ndarray) -> np.ndarray:
    """Vectorized normalized-Euclidean similarity: 1/(1+d_norm)."""
    if batch.ndim == 1:
        batch = batch.reshape(1, -1)

    all_vals = np.concatenate([model.reshape(1, -1), batch], axis=0)
    scale = max(float(np.std(all_vals)), 1e-15)

    diffs = batch - model
    dists = np.sqrt(np.sum(diffs ** 2, axis=1))
    normed = dists / (scale * np.sqrt(model.shape[0]))
    return 1.0 / (1.0 + normed)


_KERNELS = {
    "pearson": _batch_pearson,
    "cosine": _batch_cosine,
    "euclidean": _batch_euclidean,
}


# ---------------------------------------------------------------------------
# Public class
# ---------------------------------------------------------------------------

class ApproximateMatcher:
    """Vectorized approximate matching engine.

    All upstream vectors of the same length are pre-stacked into a numpy
    matrix during ``index_upstream()``, so each ``match()`` call performs
    one kernel invocation per length bucket — no Python-level per-vector loops.
    """

    def __init__(self, config: TraceConfig):
        self.config = config
        self._kernel = _KERNELS.get(config.similarity_metric, _batch_pearson)
        self._min_sim = config.min_similarity
        self._len_tol = config.length_tolerance_pct / 100.0
        # Raw vector lists (for metadata), grouped by length
        self._by_length: dict[int, list[TracingVector]] = {}
        # Pre-stacked numpy matrices, grouped by length: (N, L)
        self._arrays: dict[int, np.ndarray] = {}

    def index_upstream(self, vectors: list[TracingVector]) -> None:
        for v in vectors:
            self._by_length.setdefault(v.length, []).append(v)
        # Pre-stack into numpy arrays for fast batch operations
        self._arrays = {
            length: np.array([v.values for v in vecs])
            for length, vecs in self._by_length.items()
        }

    # ------------------------------------------------------------------ #
    # Matching
    # ------------------------------------------------------------------ #

    def match(
        self,
        model_vec: TracingVector,
        exclude: set[str] | None = None,
    ) -> list[VectorMatch]:
        """Find top-N approximate matches for *model_vec*."""
        if not self._by_length:
            return []

        exclude = exclude or set()
        model_arr = np.array(model_vec.values, dtype=np.float64)
        model_len = model_vec.length

        min_up = max(self.config.min_vector_length, int(model_len * (1.0 - self._len_tol)))
        max_up_slide = int(model_len * (1.0 + self._len_tol)) * 2

        candidates: list[tuple[float, TracingVector, str]] = []

        for up_len in self._by_length:
            vecs = self._by_length[up_len]
            batch = self._arrays[up_len]

            # Direction filter — build a boolean mask
            if self.config.direction_sensitive:
                mask = np.array([v.direction == model_vec.direction for v in vecs])
                if not np.any(mask):
                    continue
                vecs = [v for v, m in zip(vecs, mask) if m]
                batch = batch[mask]

            # Exclude already-matched
            if exclude:
                keep = np.array([
                    f"{v.file}|{v.sheet}|{v.cell_range}" not in exclude
                    for v in vecs
                ])
                if not np.any(keep):
                    continue
                vecs = [v for v, k in zip(vecs, keep) if k]
                batch = batch[keep]

            n_vecs = len(vecs)
            if n_vecs == 0:
                continue

            if up_len == model_len:
                # ── Same length: single batch kernel call ──
                sims = self._kernel(model_arr, batch)
                above = sims >= self._min_sim
                for i in np.where(above)[0]:
                    candidates.append((float(sims[i]), vecs[i], vecs[i].cell_range))

            elif up_len > model_len and up_len <= max_up_slide:
                # ── Upstream longer: 2D sliding window, one kernel call ──
                n_win = up_len - model_len + 1
                total_els = n_vecs * n_win * model_len

                if total_els <= _MAX_BATCH_ELEMENTS:
                    # sliding_window_view on 2D: (N, up_len) -> (N, n_win, model_len)
                    windows_3d = np.lib.stride_tricks.sliding_window_view(
                        batch, model_len, axis=1,
                    )
                    # Flatten to (N*n_win, model_len) — needs contiguous copy
                    flat = windows_3d.reshape(-1, model_len).copy()
                    sims = self._kernel(model_arr, flat)          # (N*n_win,)
                    sims_2d = sims.reshape(n_vecs, n_win)         # (N, n_win)
                    best_offsets = np.argmax(sims_2d, axis=1)     # (N,)
                    best_sims = sims_2d[np.arange(n_vecs), best_offsets]

                    above = best_sims >= self._min_sim
                    for i in np.where(above)[0]:
                        sub_range = compute_sub_range(vecs[i], int(best_offsets[i]), model_len)
                        candidates.append((float(best_sims[i]), vecs[i], sub_range))
                else:
                    # Fallback: per-vector to avoid memory explosion
                    for vi, v in enumerate(vecs):
                        windows = np.lib.stride_tricks.sliding_window_view(
                            batch[vi], model_len,
                        )
                        s = self._kernel(model_arr, windows)
                        bj = int(np.argmax(s))
                        bs = float(s[bj])
                        if bs >= self._min_sim:
                            candidates.append((bs, v, compute_sub_range(v, bj, model_len)))

            elif min_up <= up_len < model_len:
                # ── Upstream shorter: slide upstream over model ──
                # For each of the (model_len - up_len + 1) model windows,
                # batch-compare with all upstream vectors, take per-vector max.
                n_win = model_len - up_len + 1
                model_windows = np.lib.stride_tricks.sliding_window_view(
                    model_arr, up_len,
                )  # (n_win, up_len)

                # Stack kernel calls: (n_win,) iterations, each returns (N,)
                # n_win is tiny (usually 1-5), so the loop is negligible
                all_sims = np.empty((n_win, n_vecs), dtype=np.float64)
                for wi in range(n_win):
                    all_sims[wi] = self._kernel(model_windows[wi], batch)

                best_sims = np.max(all_sims, axis=0)  # (N,)
                above = best_sims >= self._min_sim
                for i in np.where(above)[0]:
                    candidates.append((float(best_sims[i]), vecs[i], vecs[i].cell_range))

        # Sort descending, take top-N
        candidates.sort(key=lambda x: -x[0])
        top = candidates[: self.config.top_n]

        return [
            VectorMatch(
                model_sheet=model_vec.sheet,
                model_range=model_vec.cell_range,
                model_direction=model_vec.direction,
                model_length=model_vec.length,
                model_sample=list(model_vec.values[:5]),
                match_rank=rank,
                match_type="approximate",
                similarity=round(float(sim), 6),
                upstream_file=upstream.file,
                upstream_sheet=upstream.sheet,
                upstream_range=upstream.cell_range,
                upstream_direction=upstream.direction,
                upstream_length=upstream.length,
                upstream_sample=list(upstream.values[:5]),
                upstream_matched_range=matched_range,
            )
            for rank, (sim, upstream, matched_range) in enumerate(top, 1)
        ]
