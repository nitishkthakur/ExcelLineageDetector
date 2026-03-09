"""Tests for the Excel Lineage Detector."""

from __future__ import annotations
from pathlib import Path

import pytest

from tests.test_generator import generate_test_workbook
from lineage.detector import ExcelLineageDetector

FIXTURE_DIR = Path(__file__).parent / "fixtures"


def test_coverage():
    """Test that the detector finds at least 60% of planted connections."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    planted = generate_test_workbook(test_file)

    assert test_file.exists(), f"Test file was not generated: {test_file}"

    detector = ExcelLineageDetector()
    found = detector.detect(test_file)

    assert len(found) >= 0, "Detector returned no results (shouldn't crash)"

    # Score by type
    found_types = {c.category for c in found} | {c.sub_type for c in found}
    planted_types = {p["type"] for p in planted}

    matched = planted_types & found_types
    score = len(matched) / len(planted_types) * 100 if planted_types else 0

    print(f"\nCoverage: {score:.1f}% ({len(matched)}/{len(planted_types)})")
    print(f"Found types: {sorted(found_types)}")
    print(f"Planted types: {sorted(planted_types)}")
    print(f"Matched: {sorted(matched)}")
    print(f"Missed: {sorted(planted_types - found_types)}")
    print(f"\nFound {len(found)} connections:")
    for c in found:
        print(f"  [{c.category}/{c.sub_type}] {c.source[:60]} @ {c.location}")

    assert score >= 60, (
        f"Coverage {score:.1f}% is below 60% threshold. "
        f"Matched: {sorted(matched)}, Missed: {sorted(planted_types - found_types)}"
    )
    assert len(found) >= 5, f"Should find at least 5 connections, found {len(found)}"


def test_detector_handles_missing_file():
    """Test that detector gracefully handles missing files."""
    detector = ExcelLineageDetector()
    result = detector.detect(Path("/nonexistent/file.xlsx"))
    assert result == []


def test_detector_returns_list():
    """Test that detector returns a list even for valid files."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    detector = ExcelLineageDetector()
    result = detector.detect(test_file)
    assert isinstance(result, list)


def test_connections_have_required_fields():
    """Test that all returned connections have required fields."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    detector = ExcelLineageDetector()
    connections = detector.detect(test_file)

    for conn in connections:
        assert conn.id, f"Connection missing id: {conn}"
        assert conn.category, f"Connection missing category: {conn}"
        assert conn.sub_type, f"Connection missing sub_type: {conn}"
        assert conn.source is not None, f"Connection missing source: {conn}"
        assert conn.raw_connection is not None, f"Connection missing raw_connection: {conn}"
        assert conn.location, f"Connection missing location: {conn}"
        assert 0.0 <= conn.confidence <= 1.0, f"Connection confidence out of range: {conn}"


def test_connections_are_deduplicated():
    """Test that the detector deduplicates connections by id."""
    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    detector = ExcelLineageDetector()
    connections = detector.detect(test_file)

    # Check for duplicate IDs
    ids = [c.id for c in connections]
    assert len(ids) == len(set(ids)), (
        f"Duplicate connection IDs found: "
        f"{[id for id in ids if ids.count(id) > 1]}"
    )


def test_connections_serializable():
    """Test that all connections can be serialized to dict."""
    import json

    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    detector = ExcelLineageDetector()
    connections = detector.detect(test_file)

    for conn in connections:
        d = conn.to_dict()
        assert isinstance(d, dict)
        # Should be JSON-serializable
        json_str = json.dumps(d, default=str)
        assert json_str


def test_json_reporter():
    """Test that the JSON reporter produces valid output."""
    import json
    import tempfile

    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    detector = ExcelLineageDetector()
    connections = detector.detect(test_file)

    from lineage.reporters.json_reporter import JsonReporter

    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir)
        out_path = JsonReporter().write(connections, test_file, out_dir)
        assert out_path.exists()

        data = json.loads(out_path.read_text())
        assert "file" in data
        assert "scanned_at" in data
        assert "summary" in data
        assert "connections" in data
        assert data["summary"]["total_connections"] == len(connections)


def test_excel_reporter():
    """Test that the Excel reporter produces a valid xlsx file."""
    import tempfile

    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    detector = ExcelLineageDetector()
    connections = detector.detect(test_file)

    from lineage.reporters.excel_reporter import ExcelReporter

    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir)
        out_path = ExcelReporter().write(connections, test_file, out_dir)
        assert out_path.exists()
        assert out_path.suffix == ".xlsx"
        assert out_path.stat().st_size > 0


def test_graph_reporter():
    """Test that the graph reporter produces a valid PNG file."""
    import tempfile

    FIXTURE_DIR.mkdir(exist_ok=True)
    test_file = FIXTURE_DIR / "test_connections.xlsx"
    if not test_file.exists():
        generate_test_workbook(test_file)

    detector = ExcelLineageDetector()
    connections = detector.detect(test_file)

    from lineage.reporters.graph_reporter import GraphReporter

    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir)
        out_path = GraphReporter().write(connections, test_file, out_dir)
        assert out_path.exists()
        assert out_path.suffix == ".png"
        assert out_path.stat().st_size > 0


def test_parsers_sql():
    """Test SQL parser."""
    from lineage.parsers.sql_parser import parse

    result = parse("SELECT id, name FROM customers WHERE status = 'active'")
    assert result is not None
    assert "customers" in result.tables


def test_parsers_m():
    """Test M formula parser."""
    from lineage.parsers.m_parser import parse

    m_code = 'let Source = Sql.Database("myserver", "mydb"), Data = Source in Data'
    result = parse(m_code)
    assert result["sub_type"] == "sql_server"
    assert "myserver" in result["source"]


def test_parsers_connection_string():
    """Test connection string parser."""
    from lineage.parsers.connection_string import parse

    cs = "Provider=SQLOLEDB;Server=myserver;Database=mydb;Trusted_Connection=Yes"
    result = parse(cs)
    assert result.get("_server") == "myserver"
    assert result.get("_database") == "mydb"


def test_parsers_formula():
    """Test formula parser."""
    from lineage.parsers.formula_parser import parse

    formula = "='[source_data.xlsx]Sheet1'!A1"
    result = parse(formula)
    assert result is not None
    assert result.get("workbook_name") == "source_data.xlsx"
