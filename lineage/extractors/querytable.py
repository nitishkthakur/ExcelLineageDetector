"""Extractor for Query Tables."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection
from lineage.parsers.sql_parser import parse as parse_sql


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class QueryTableExtractor(BaseExtractor):
    """Extracts data connections from Query Tables."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            qt_files = [n for n in zip_file.namelist()
                        if re.match(r"xl/queryTables/queryTable\d+\.xml$", n)]

            if not qt_files:
                self.log.debug("No query table files found")
                return connections

            # Load connection info for cross-referencing
            conn_map = self._load_connections(zip_file)

            for qt_file in qt_files:
                try:
                    found = self._extract_from_query_table(zip_file, qt_file, conn_map)
                    connections.extend(found)
                except Exception as e:
                    self.log.warning(f"Failed to extract from {qt_file}: {e}")

        except Exception as e:
            self.log.error(f"QueryTableExtractor failed: {e}", exc_info=True)

        return connections

    def _load_connections(self, zip_file: zipfile.ZipFile) -> dict[str, dict]:
        """Load connections from connections.xml for cross-referencing."""
        conn_map = {}
        root = self._read_xml(zip_file, "xl/connections.xml")
        if root is None:
            return conn_map

        conn_els = root.findall(f"{{{NS}}}connection")
        if not conn_els:
            conn_els = root.findall("connection")
        if not conn_els:
            conn_els = root.findall("{*}connection")

        for conn_el in conn_els:
            conn_id = conn_el.get("id", "")
            conn_name = conn_el.get("name", "")
            conn_type = conn_el.get("type", "")

            db_pr = conn_el.find(f"{{{NS}}}dbPr") or conn_el.find("dbPr") or conn_el.find("{*}dbPr")
            cs = db_pr.get("connection", "") if db_pr is not None else ""
            cmd = db_pr.get("command", "") if db_pr is not None else ""

            conn_map[conn_id] = {
                "id": conn_id,
                "name": conn_name,
                "type": conn_type,
                "connection_string": cs,
                "command": cmd,
            }

        return conn_map

    def _extract_from_query_table(
        self,
        zip_file: zipfile.ZipFile,
        qt_file: str,
        conn_map: dict,
    ) -> list[DataConnection]:
        """Extract connections from a single query table file."""
        results = []

        root = self._read_xml(zip_file, qt_file)
        if root is None:
            return results

        # Get query table attributes
        qt_name = root.get("name", "QueryTable")
        connection_id = root.get("connectionId", "")
        auto_format_id = root.get("autoFormatId", "")

        # Get sheet name/location
        sheet_ref = self._get_query_table_location(zip_file, qt_file)
        location = f"{sheet_ref}:QueryTable:{qt_name}" if sheet_ref else f"QueryTable:{qt_name}"

        # Look for SQL override in refresh
        qt_refresh = root.find(f"{{{NS}}}queryTableRefresh")
        if qt_refresh is None:
            qt_refresh = root.find("queryTableRefresh")
        if qt_refresh is None:
            qt_refresh = root.find("{*}queryTableRefresh")

        if qt_refresh is not None:
            # Check for SQL in table fields or range
            sql_el = qt_refresh.find(f"{{{NS}}}queryTableFields")
            if sql_el is not None:
                # Extract field names as hint
                pass

        # Look up connection
        conn_info = conn_map.get(connection_id, {})
        raw_connection = conn_info.get("connection_string", "") or conn_info.get("name", "")
        conn_name = conn_info.get("name", qt_name)
        conn_type = conn_info.get("type", "")
        query_text = conn_info.get("command", "") or None
        parsed_query = parse_sql(query_text) if query_text else None

        # Map type to sub_type
        type_map = {
            "1": "odbc", "2": "dao", "3": "web",
            "4": "oledb", "5": "text", "6": "ado",
        }
        sub_type = type_map.get(conn_type, "querytable")

        if not raw_connection:
            raw_connection = qt_name

        conn = DataConnection(
            id=DataConnection.make_id("database", raw_connection, location),
            category="database",
            sub_type=sub_type,
            source=conn_name or qt_name,
            raw_connection=raw_connection,
            location=location,
            query_text=query_text,
            parsed_query=parsed_query,
            metadata={
                "query_table_name": qt_name,
                "connection_id": connection_id,
                "connection_name": conn_name,
            },
        )
        results.append(conn)

        return results

    def _get_query_table_location(self, zip_file: zipfile.ZipFile, qt_file: str) -> str:
        """Get the sheet location of this query table via relationships."""
        try:
            qt_rels_file = qt_file.replace(
                "xl/queryTables/", "xl/queryTables/_rels/"
            ).replace(".xml", ".xml.rels")

            if qt_rels_file in zip_file.namelist():
                rels = self._read_rels(zip_file, qt_rels_file)
                for rid, rel_info in rels.items():
                    target = rel_info["target"]
                    if "worksheets" in target:
                        match = re.search(r"sheet(\d+)\.xml", target)
                        if match:
                            return f"Sheet{match.group(1)}"
        except Exception:
            pass
        return ""
