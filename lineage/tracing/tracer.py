"""Upstream tracer — orchestrates scanning, matching, and result collection."""
from __future__ import annotations

import os
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from lineage.tracing.config import TraceConfig
from lineage.tracing.models import TracingVector, VectorMatch


# ---------------------------------------------------------------------------
# Top-level worker (must be picklable for ProcessPoolExecutor)
# ---------------------------------------------------------------------------

def _scan_worker(args: tuple[str, int]) -> list[dict]:
    """Scan one upstream file in a worker process.

    Returns dicts (not TracingVector) for pickle compatibility,
    though dataclasses *do* pickle fine in CPython 3.10+.
    """
    path_str, min_len = args
    from lineage.tracing.scanner import scan_upstream_file
    vectors = scan_upstream_file(Path(path_str), min_len)
    return [
        {
            "file": v.file,
            "sheet": v.sheet,
            "cell_range": v.cell_range,
            "direction": v.direction,
            "length": v.length,
            "start_cell": v.start_cell,
            "end_cell": v.end_cell,
            "values": v.values,
        }
        for v in vectors
    ]


class UpstreamTracer:
    """High-level orchestrator for upstream vector tracing."""

    def __init__(self, config: TraceConfig | None = None, verbose: bool = False):
        self.config = config or TraceConfig()
        self.verbose = verbose

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    def trace(
        self,
        model_path: Path,
        sheet_name: str,
        upstream_paths: list[Path],
    ) -> tuple[list[VectorMatch], list[TracingVector]]:
        """Run the full tracing pipeline.

        Returns
        -------
        matches : list[VectorMatch]
            All matches (exact + approximate), ordered by model vector then rank.
        unmatched : list[TracingVector]
            Model vectors that had zero matches.
        """
        t0 = time.perf_counter()

        # 1. Scan model sheet
        if self.verbose:
            print(f"Scanning model sheet '{sheet_name}' in {model_path.name} ...")
        from lineage.tracing.scanner import scan_model_sheet
        model_vectors = scan_model_sheet(
            model_path, sheet_name, self.config.min_vector_length,
        )
        if self.verbose:
            print(f"  Found {len(model_vectors)} hardcoded vectors")

        if not model_vectors:
            return [], []

        # 2. Scan upstream files (parallel)
        upstream_vectors = self._scan_upstream_parallel(upstream_paths)
        if self.verbose:
            print(f"  Total upstream vectors: {len(upstream_vectors)}")

        if not upstream_vectors:
            return [], list(model_vectors)

        # 3. Build matchers
        from lineage.tracing.exact_matcher import ExactMatcher
        from lineage.tracing.approx_matcher import ApproximateMatcher

        exact_matcher = ExactMatcher(self.config) if self.config.exact else None
        approx_matcher = ApproximateMatcher(self.config) if self.config.approximate else None

        if exact_matcher:
            exact_matcher.index_upstream(upstream_vectors)
        if approx_matcher:
            approx_matcher.index_upstream(upstream_vectors)

        # 4. Match each model vector
        all_matches: list[VectorMatch] = []
        unmatched: list[TracingVector] = []

        for mi, mvec in enumerate(model_vectors):
            if self.verbose and (mi + 1) % 100 == 0:
                print(f"  Matching vector {mi + 1}/{len(model_vectors)} ...")

            vec_matches: list[VectorMatch] = []
            exact_keys: set[str] = set()

            # Exact
            if exact_matcher:
                exact_hits = exact_matcher.match(mvec)
                for m in exact_hits:
                    exact_keys.add(f"{m.upstream_file}|{m.upstream_sheet}|{m.upstream_range}")
                vec_matches.extend(exact_hits)

            # Approximate (excluding already-exact-matched upstream vectors)
            if approx_matcher:
                approx_hits = approx_matcher.match(mvec, exclude=exact_keys)
                vec_matches.extend(approx_hits)

            if vec_matches:
                # Assign ranks: exact first, then approximate by similarity
                for rank, m in enumerate(vec_matches, 1):
                    m.match_rank = rank
                all_matches.extend(vec_matches)
            else:
                unmatched.append(mvec)

        elapsed = time.perf_counter() - t0
        if self.verbose:
            n_exact = sum(1 for m in all_matches if m.match_type.startswith("exact"))
            n_approx = sum(1 for m in all_matches if m.match_type == "approximate")
            print(f"  Done: {n_exact} exact + {n_approx} approximate matches "
                  f"for {len(model_vectors)} vectors ({elapsed:.2f}s)")
            if unmatched:
                print(f"  {len(unmatched)} vectors had no matches")

        return all_matches, unmatched

    # ------------------------------------------------------------------ #
    # Parallel upstream scanning
    # ------------------------------------------------------------------ #

    def _scan_upstream_parallel(self, paths: list[Path]) -> list[TracingVector]:
        """Scan all upstream files using ProcessPoolExecutor."""
        if not paths:
            return []

        # Filter to XLSX/XLSM only
        valid: list[Path] = []
        for p in paths:
            suffix = p.suffix.lower()
            if suffix in (".xlsx", ".xlsm"):
                valid.append(p)
            elif self.verbose:
                print(f"  Skipping {p.name} (only .xlsx/.xlsm supported for tracing)")

        if not valid:
            return []

        min_len = self.config.min_vector_length
        workers = self.config.max_workers or min(os.cpu_count() or 4, len(valid))

        if self.verbose:
            print(f"Scanning {len(valid)} upstream file(s) with {workers} worker(s) ...")

        t0 = time.perf_counter()
        all_vectors: list[TracingVector] = []

        # Use ProcessPoolExecutor for true parallelism (lxml is CPU-bound)
        args_list = [(str(p), min_len) for p in valid]

        try:
            with ProcessPoolExecutor(max_workers=workers) as pool:
                futures = {pool.submit(_scan_worker, a): a[0] for a in args_list}
                for future in as_completed(futures):
                    fname = Path(futures[future]).name
                    try:
                        dicts = future.result()
                        vectors = [
                            TracingVector(
                                file=d["file"],
                                sheet=d["sheet"],
                                cell_range=d["cell_range"],
                                direction=d["direction"],
                                length=d["length"],
                                start_cell=d["start_cell"],
                                end_cell=d["end_cell"],
                                values=tuple(d["values"]),
                            )
                            for d in dicts
                        ]
                        all_vectors.extend(vectors)
                        if self.verbose:
                            print(f"    {fname}: {len(vectors)} vectors")
                    except Exception as e:
                        if self.verbose:
                            print(f"    {fname}: ERROR — {e}")
        except Exception:
            # Fallback to sequential if multiprocessing fails
            if self.verbose:
                print("  Parallel scan failed, falling back to sequential ...")
            from lineage.tracing.scanner import scan_upstream_file
            for p in valid:
                vectors = scan_upstream_file(p, min_len)
                all_vectors.extend(vectors)
                if self.verbose:
                    print(f"    {p.name}: {len(vectors)} vectors")

        if self.verbose:
            elapsed = time.perf_counter() - t0
            print(f"  Upstream scan: {len(all_vectors)} vectors in {elapsed:.2f}s")

        return all_vectors
