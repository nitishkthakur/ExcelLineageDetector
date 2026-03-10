"""Extractor for Power Query M code."""

from __future__ import annotations
import json
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection
from lineage.parsers.m_parser import parse as parse_m, parse_all as parse_m_all


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class PowerQueryExtractor(BaseExtractor):
    """Extracts Power Query M formulas from Excel files."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            connections.extend(self._extract_from_custom_xml(zip_file))
        except Exception as e:
            self.log.error(f"PowerQueryExtractor (customXml) failed: {e}", exc_info=True)

        try:
            connections.extend(self._extract_from_connections(zip_file))
        except Exception as e:
            self.log.error(f"PowerQueryExtractor (connections) failed: {e}", exc_info=True)

        return connections

    def _extract_from_custom_xml(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract Power Query from customXml items."""
        results = []
        names = zip_file.namelist()

        for name in names:
            if not name.startswith("xl/customXml/") or not name.endswith(".xml"):
                continue
            try:
                data = zip_file.read(name)
                results.extend(self._parse_custom_xml_data(data, name))
            except Exception as e:
                self.log.warning(f"Failed to read {name}: {e}")

        return results

    def _parse_custom_xml_data(self, data: bytes, source_name: str) -> list[DataConnection]:
        """Parse a customXml data item for Power Query content."""
        results = []

        # Try to decode as text
        try:
            text = data.decode("utf-8", errors="replace")
        except Exception:
            return results

        # Look for JSON content (Power Query stores queries as JSON)
        # Pattern: {"Queries":[...]} or similar
        json_matches = re.findall(r'\{[^{}]*"(?:Formula|Queries|Name)"[^{}]*\}', text, re.DOTALL)

        # Try to parse whole thing as JSON (sometimes it's wrapped in XML)
        json_start = text.find('{')
        json_end = text.rfind('}')
        if json_start != -1 and json_end != -1:
            json_str = text[json_start:json_end + 1]
            try:
                data_obj = json.loads(json_str)
                queries = None
                if isinstance(data_obj, dict):
                    # Look for queries in various formats
                    queries = data_obj.get("Queries", data_obj.get("queries", None))
                    if queries is None and "Formula" in data_obj:
                        queries = [data_obj]
                    if queries is None:
                        # Recurse into nested objects
                        for v in data_obj.values():
                            if isinstance(v, list) and v and isinstance(v[0], dict):
                                if "Formula" in v[0] or "Name" in v[0]:
                                    queries = v
                                    break

                if queries:
                    for query in queries:
                        if not isinstance(query, dict):
                            continue
                        conns = self._make_query_connections(query, source_name)
                        results.extend(conns)
            except json.JSONDecodeError:
                pass

        # Also try to find individual M formula blocks in XML
        formula_pattern = re.compile(r'<[^>]*Formula[^>]*>([^<]+)</[^>]*Formula[^>]*>', re.IGNORECASE)
        name_pattern = re.compile(r'<[^>]*Name[^>]*>([^<]+)</[^>]*Name[^>]*>', re.IGNORECASE)

        formula_matches = formula_pattern.findall(text)
        name_matches = name_pattern.findall(text)

        for i, formula in enumerate(formula_matches):
            formula = formula.strip()
            if len(formula) < 5:
                continue
            name = name_matches[i] if i < len(name_matches) else f"Query_{i}"
            all_parsed = parse_m_all(formula)
            if not all_parsed:
                all_parsed = [{"sub_type": "m_formula", "source": name, "details": {}}]

            for j, parsed in enumerate(all_parsed):
                raw_conn = parsed.get("source", "") or formula[:100]
                id_key = f"{formula[:50]}:{j}" if j > 0 else formula[:50]
                conn = DataConnection(
                    id=DataConnection.make_id("powerquery", id_key, source_name),
                    category="powerquery",
                    sub_type=parsed.get("sub_type", "m_formula"),
                    source=parsed.get("source", name) or name,
                    raw_connection=raw_conn,
                    location=source_name,
                    query_text=formula if j == 0 else None,
                    metadata={
                        "query_name": name,
                        "details": parsed.get("details", {}),
                    },
                )
                results.append(conn)

        return results

    def _make_query_connection(self, query: dict, source_name: str) -> DataConnection | None:
        """Create a DataConnection from a parsed Power Query query object.

        Returns the first connection; additional sources in the same M script
        are emitted via _make_query_connections().
        """
        conns = self._make_query_connections(query, source_name)
        return conns[0] if conns else None

    def _make_query_connections(
        self, query: dict, source_name: str
    ) -> list[DataConnection]:
        """Create DataConnection(s) from a Power Query object - one per data source."""
        formula = query.get("Formula", query.get("formula", ""))
        name = query.get("Name", query.get("name", "UnnamedQuery"))
        description = query.get("Description", "")

        if not formula:
            return []

        all_parsed = parse_m_all(formula)
        if not all_parsed:
            all_parsed = [{"sub_type": "m_formula", "source": name, "details": {}}]

        results = []
        for i, parsed in enumerate(all_parsed):
            raw_conn = parsed.get("source", "") or formula[:100]
            id_key = f"{name}:{i}" if i > 0 else name
            conn = DataConnection(
                id=DataConnection.make_id("powerquery", id_key, source_name),
                category="powerquery",
                sub_type=parsed.get("sub_type", "m_formula"),
                source=parsed.get("source", name) or name,
                raw_connection=raw_conn,
                location=source_name,
                query_text=formula if i == 0 else None,
                metadata={
                    "query_name": name,
                    "description": description,
                    "details": parsed.get("details", {}),
                },
            )
            results.append(conn)
        return results

    def _extract_from_connections(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract Power Query embedded in connections.xml (Query - prefix pattern)."""
        results = []
        root = self._read_xml(zip_file, "xl/connections.xml")
        if root is None:
            return results

        connections_el = root.findall(f"{{{NS}}}connection")
        if not connections_el:
            connections_el = root.findall("connection")
        if not connections_el:
            connections_el = root.findall(".//{*}connection")

        for conn_el in connections_el:
            name = conn_el.get("name", "")
            # Power Query connections start with "Query - "
            if not name.startswith("Query - ") and not name.startswith("Query_"):
                continue

            query_name = name.replace("Query - ", "").replace("Query_", "")

            # Look for embedded M formula in dbPr or oleDb
            db_pr = conn_el.find(f"{{{NS}}}dbPr")
            if db_pr is None:
                db_pr = conn_el.find("dbPr")
            if db_pr is None:
                db_pr = conn_el.find("{*}dbPr")

            raw_conn = ""
            formula = ""
            if db_pr is not None:
                raw_conn = db_pr.get("connection", "")
                formula = db_pr.get("command", "")

            if not formula and not raw_conn:
                raw_conn = name

            parsed = parse_m(formula) if formula else {"sub_type": "powerquery", "source": raw_conn, "details": {}}

            conn = DataConnection(
                id=DataConnection.make_id("powerquery", name, "connections.xml"),
                category="powerquery",
                sub_type=parsed.get("sub_type", "powerquery"),
                source=parsed.get("source", query_name) or query_name,
                raw_connection=raw_conn or name,
                location="connections.xml",
                query_text=formula or None,
                metadata={"query_name": query_name, "details": parsed.get("details", {})},
            )
            results.append(conn)

        return results
