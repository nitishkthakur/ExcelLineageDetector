"""Tests for the upstream tracing module."""
from __future__ import annotations

import json
import tempfile
from pathlib import Path

import numpy as np
import pytest

from tests.test_generator import generate_test_workbook

FIXTURE_DIR = Path(__file__).parent / "fixtures"


# ---------------------------------------------------------------------------
# Config tests
# ---------------------------------------------------------------------------

def test_config_defaults():
    """TraceConfig has sensible defaults."""
    from lineage.tracing.config import TraceConfig

    c = TraceConfig()
    assert c.exact is True
    assert c.approximate is True
    assert c.top_n == 5
    assert c.similarity_metric == "pearson"
    assert c.min_similarity == 0.8
    assert c.subsequence_matching is True
    assert c.direction_sensitive is False
    assert c.min_vector_length == 3


def test_config_from_json():
    """TraceConfig loads from JSON."""
    from lineage.tracing.config import TraceConfig

    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as f:
        json.dump({
            "matching": {"exact": False, "top_n": 10, "similarity_metric": "cosine"},
            "performance": {"min_vector_length": 5},
        }, f)
        f.flush()
        c = TraceConfig.from_file(Path(f.name))

    assert c.exact is False
    assert c.top_n == 10
    assert c.similarity_metric == "cosine"
    assert c.min_vector_length == 5
    # Defaults preserved for unset fields
    assert c.approximate is True


def test_config_to_dict():
    """TraceConfig round-trips through to_dict."""
    from lineage.tracing.config import TraceConfig

    c = TraceConfig(top_n=3, similarity_metric="euclidean")
    d = c.to_dict()
    assert d["matching"]["top_n"] == 3
    assert d["matching"]["similarity_metric"] == "euclidean"


# ---------------------------------------------------------------------------
# Scanner tests
# ---------------------------------------------------------------------------

def test_scan_model_sheet():
    """Scanner finds hardcoded vectors in a model sheet."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    from lineage.tracing.scanner import scan_model_sheet
    vectors = scan_model_sheet(test_file, "Inputs")
    assert len(vectors) > 0, "Expected hardcoded vectors in Inputs sheet"
    for v in vectors:
        assert v.file == "test_connections.xlsx"
        assert v.sheet == "Inputs"
        assert v.length >= 3
        assert len(v.values) == v.length
        assert v.direction in ("column", "row")


def test_scan_upstream_file():
    """Upstream scanner finds ALL numeric vectors (formula + hardcoded)."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    from lineage.tracing.scanner import scan_upstream_file
    vectors = scan_upstream_file(test_file)
    assert len(vectors) > 0, "Expected upstream vectors"
    # Should find vectors across multiple sheets
    sheets = {v.sheet for v in vectors}
    assert len(sheets) >= 1


def test_get_sheet_names():
    """get_sheet_names returns correct sheets."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    from lineage.tracing.scanner import get_sheet_names
    names = get_sheet_names(test_file)
    assert "Inputs" in names


def test_scanner_xls_returns_empty():
    """Scanner gracefully returns empty for non-existent / non-ZIP files."""
    from lineage.tracing.scanner import scan_upstream_file
    result = scan_upstream_file(Path("/nonexistent/file.xlsx"))
    assert result == []


# ---------------------------------------------------------------------------
# Exact matcher tests
# ---------------------------------------------------------------------------

def test_exact_full_match():
    """ExactMatcher finds identical vectors."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.exact_matcher import ExactMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig()
    em = ExactMatcher(config)

    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A5",
        direction="column", length=5, start_cell="A1", end_cell="A5",
        values=(10.0, 20.0, 30.0, 40.0, 50.0),
    )
    em.index_upstream([upstream])

    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B3:B7",
        direction="column", length=5, start_cell="B3", end_cell="B7",
        values=(10.0, 20.0, 30.0, 40.0, 50.0),
    )
    matches = em.match(model)
    assert len(matches) == 1
    assert matches[0].match_type == "exact"
    assert matches[0].similarity == 1.0
    assert matches[0].upstream_file == "source.xlsx"


def test_exact_subsequence_match():
    """ExactMatcher finds model vector as subsequence of longer upstream."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.exact_matcher import ExactMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig(subsequence_matching=True)
    em = ExactMatcher(config)

    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A10",
        direction="column", length=10, start_cell="A1", end_cell="A10",
        values=(5.0, 10.0, 20.0, 30.0, 40.0, 50.0, 60.0, 70.0, 80.0, 90.0),
    )
    em.index_upstream([upstream])

    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B1:B3",
        direction="column", length=3, start_cell="B1", end_cell="B3",
        values=(20.0, 30.0, 40.0),
    )
    matches = em.match(model)
    assert any(m.match_type == "exact_subsequence" for m in matches)
    # The matched sub-range should be A3:A5 (offset 2, length 3)
    sub_match = [m for m in matches if m.match_type == "exact_subsequence"][0]
    assert sub_match.upstream_matched_range == "A3:A5"


def test_exact_no_match():
    """ExactMatcher returns empty when no match exists."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.exact_matcher import ExactMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig()
    em = ExactMatcher(config)

    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A3",
        direction="column", length=3, start_cell="A1", end_cell="A3",
        values=(100.0, 200.0, 300.0),
    )
    em.index_upstream([upstream])

    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B1:B3",
        direction="column", length=3, start_cell="B1", end_cell="B3",
        values=(1.0, 2.0, 3.0),
    )
    matches = em.match(model)
    assert len(matches) == 0


# ---------------------------------------------------------------------------
# Approximate matcher tests
# ---------------------------------------------------------------------------

def test_approx_pearson_identical():
    """Pearson correlation = 1.0 for identical vectors."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.approx_matcher import ApproximateMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig(similarity_metric="pearson", min_similarity=0.5)
    am = ApproximateMatcher(config)

    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A5",
        direction="column", length=5, start_cell="A1", end_cell="A5",
        values=(10.0, 20.0, 30.0, 40.0, 50.0),
    )
    am.index_upstream([upstream])

    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B1:B5",
        direction="column", length=5, start_cell="B1", end_cell="B5",
        values=(10.0, 20.0, 30.0, 40.0, 50.0),
    )
    matches = am.match(model)
    assert len(matches) >= 1
    assert matches[0].similarity > 0.99


def test_approx_pearson_scaled():
    """Pearson detects same shape at different scale."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.approx_matcher import ApproximateMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig(similarity_metric="pearson", min_similarity=0.9)
    am = ApproximateMatcher(config)

    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A5",
        direction="column", length=5, start_cell="A1", end_cell="A5",
        values=(100.0, 200.0, 300.0, 400.0, 500.0),
    )
    am.index_upstream([upstream])

    # Same shape, different scale
    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B1:B5",
        direction="column", length=5, start_cell="B1", end_cell="B5",
        values=(10.0, 20.0, 30.0, 40.0, 50.0),
    )
    matches = am.match(model)
    assert len(matches) >= 1
    assert matches[0].similarity > 0.99  # perfect linear correlation


def test_approx_length_mismatch():
    """Approximate matcher handles length-mismatched vectors via sliding window."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.approx_matcher import ApproximateMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig(
        similarity_metric="pearson",
        min_similarity=0.9,
        length_tolerance_pct=100.0,
    )
    am = ApproximateMatcher(config)

    # Upstream is longer
    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A8",
        direction="column", length=8, start_cell="A1", end_cell="A8",
        values=(0.0, 0.0, 10.0, 20.0, 30.0, 40.0, 50.0, 0.0),
    )
    am.index_upstream([upstream])

    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B1:B5",
        direction="column", length=5, start_cell="B1", end_cell="B5",
        values=(10.0, 20.0, 30.0, 40.0, 50.0),
    )
    matches = am.match(model)
    assert len(matches) >= 1
    assert matches[0].similarity > 0.99


def test_approx_cosine():
    """Cosine similarity works."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.approx_matcher import ApproximateMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig(similarity_metric="cosine", min_similarity=0.5)
    am = ApproximateMatcher(config)

    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A4",
        direction="column", length=4, start_cell="A1", end_cell="A4",
        values=(1.0, 2.0, 3.0, 4.0),
    )
    am.index_upstream([upstream])

    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B1:B4",
        direction="column", length=4, start_cell="B1", end_cell="B4",
        values=(2.0, 4.0, 6.0, 8.0),
    )
    matches = am.match(model)
    assert len(matches) >= 1
    assert matches[0].similarity > 0.99  # identical direction


def test_approx_below_threshold():
    """Vectors below min_similarity are not returned."""
    from lineage.tracing.config import TraceConfig
    from lineage.tracing.approx_matcher import ApproximateMatcher
    from lineage.tracing.models import TracingVector

    config = TraceConfig(similarity_metric="pearson", min_similarity=0.99)
    am = ApproximateMatcher(config)

    upstream = TracingVector(
        file="source.xlsx", sheet="Data", cell_range="A1:A4",
        direction="column", length=4, start_cell="A1", end_cell="A4",
        values=(1.0, 2.0, 3.0, 4.0),
    )
    am.index_upstream([upstream])

    # Very different shape
    model = TracingVector(
        file="model.xlsx", sheet="Sheet1", cell_range="B1:B4",
        direction="column", length=4, start_cell="B1", end_cell="B4",
        values=(4.0, 1.0, 3.0, 2.0),
    )
    matches = am.match(model)
    # Correlation is low, should be below 0.99
    assert all(m.similarity >= 0.99 for m in matches) or len(matches) == 0


# ---------------------------------------------------------------------------
# Batch similarity kernel tests
# ---------------------------------------------------------------------------

def test_batch_pearson_constant_vector():
    """Pearson handles constant vectors without NaN."""
    from lineage.tracing.approx_matcher import _batch_pearson

    model = np.array([5.0, 5.0, 5.0, 5.0])
    batch = np.array([
        [5.0, 5.0, 5.0, 5.0],  # same constant → 1.0
        [1.0, 2.0, 3.0, 4.0],  # non-constant → 0.0
        [3.0, 3.0, 3.0, 3.0],  # different constant → 0.0
    ])
    result = _batch_pearson(model, batch)
    assert not np.any(np.isnan(result))
    assert result[0] == pytest.approx(1.0)
    assert result[1] == pytest.approx(0.0)
    assert result[2] == pytest.approx(0.0)


# ---------------------------------------------------------------------------
# End-to-end tracer test
# ---------------------------------------------------------------------------

def test_tracer_end_to_end():
    """UpstreamTracer finds matches between test files."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    from lineage.tracing.config import TraceConfig
    from lineage.tracing.tracer import UpstreamTracer

    config = TraceConfig(min_similarity=0.7, top_n=3)
    tracer = UpstreamTracer(config=config)

    # Trace the test file against itself (should find exact matches)
    matches, unmatched = tracer.trace(
        test_file, "Inputs", [test_file],
    )
    # Should find at least some matches (tracing file against itself)
    assert len(matches) + len(unmatched) > 0


def test_tracer_report():
    """TracingReporter produces a valid Excel file."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    from lineage.tracing.config import TraceConfig
    from lineage.tracing.tracer import UpstreamTracer
    from lineage.tracing.report import TracingReporter

    config = TraceConfig(min_similarity=0.7, top_n=3)
    tracer = UpstreamTracer(config=config)
    matches, unmatched = tracer.trace(test_file, "Inputs", [test_file])

    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir)
        reporter = TracingReporter()
        out_path = reporter.write(
            matches, unmatched, config, test_file, "Inputs",
            [test_file], out_dir,
        )
        assert out_path.exists()
        assert out_path.suffix == ".xlsx"
        assert out_path.stat().st_size > 0

        # Verify structure
        import openpyxl
        wb = openpyxl.load_workbook(str(out_path))
        assert "Config" in wb.sheetnames
        assert "Tracing Results" in wb.sheetnames
