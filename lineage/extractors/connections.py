"""Extractor for xl/connections.xml data connections."""

from __future__ import annotations
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection
from lineage.parsers.connection_string import parse as parse_cs, format_source_label


# Connection type int to sub_type mapping per OOXML spec
CONNECTION_TYPE_MAP = {
    "1": "odbc",
    "2": "dao",
    "3": "web",
    "4": "oledb",
    "5": "text",
    "6": "ado",
    "7": "dsp",
}

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


class ConnectionsExtractor(BaseExtractor):
    """Extracts database/web/file connections from xl/connections.xml."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            connections.extend(self._parse_connections(zip_file))
        except Exception as e:
            self.log.error(f"ConnectionsExtractor failed: {e}", exc_info=True)
        return connections

    def _parse_connections(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        results = []
        root = self._read_xml(zip_file, "xl/connections.xml")
        if root is None:
            self.log.debug("No connections.xml found")
            return results

        # Find all connection elements (handle namespace)
        connections_el = []
        # Try with namespace
        connections_el = root.findall(f"{{{NS}}}connection")
        if not connections_el:
            # Try without namespace
            connections_el = root.findall("connection")
        if not connections_el:
            # Try wildcard search
            connections_el = root.findall(".//{*}connection")

        for conn_el in connections_el:
            try:
                result = self._parse_connection_element(conn_el)
                if result:
                    results.append(result)
            except Exception as e:
                self.log.warning(f"Failed to parse connection element: {e}")

        return results

    def _parse_connection_element(self, conn_el) -> DataConnection | None:
        conn_id = conn_el.get("id", "")
        conn_name = conn_el.get("name", f"connection_{conn_id}")
        conn_type = conn_el.get("type", "0")
        refreshed_version = conn_el.get("refreshedVersion", "")
        background = conn_el.get("background", "0")
        save_password = conn_el.get("savePassword", "0")

        sub_type = CONNECTION_TYPE_MAP.get(conn_type, f"type_{conn_type}")

        raw_connection = ""
        query_text = None
        category = "database"
        source = conn_name

        # Check for dbPr (database connection)
        db_pr = self._find_child(conn_el, "dbPr")
        if db_pr is not None:
            raw_connection = db_pr.get("connection", "")
            command = db_pr.get("command", "")
            if command:
                query_text = command
            category = "database"

            # Parse connection string
            parsed_cs = parse_cs(raw_connection)
            source = format_source_label(parsed_cs) or conn_name
            if parsed_cs.get("_sub_type"):
                sub_type = parsed_cs["_sub_type"]

        # Check for webPr (web query)
        web_pr = self._find_child(conn_el, "webPr")
        if web_pr is not None:
            url = web_pr.get("url", "")
            raw_connection = url
            category = "web"
            sub_type = "web_query"
            source = url[:80] if url else conn_name

        # Check for textPr (text file)
        text_pr = self._find_child(conn_el, "textPr")
        if text_pr is not None:
            source_file = text_pr.get("sourceFile", "")
            raw_connection = source_file
            category = "file"
            sub_type = "text"
            source = source_file or conn_name

        # Check for olapPr (OLAP)
        olap_pr = self._find_child(conn_el, "olapPr")
        if olap_pr is not None:
            category = "database"
            sub_type = "olap"
            if not raw_connection:
                raw_connection = conn_name

        if not raw_connection:
            raw_connection = conn_name or f"connection_{conn_id}"

        location = "connections.xml"

        conn = DataConnection(
            id=DataConnection.make_id(category, raw_connection, location),
            category=category,
            sub_type=sub_type,
            source=source,
            raw_connection=raw_connection,
            location=location,
            query_text=query_text,
            metadata={
                "conn_id": conn_id,
                "conn_name": conn_name,
                "type": conn_type,
                "refreshed_version": refreshed_version,
                "background": background,
                "save_password": save_password,
            },
        )

        return conn

    def _find_child(self, element, tag: str):
        """Find a child element by tag name, with or without namespace."""
        # Try with namespace
        child = element.find(f"{{{NS}}}{tag}")
        if child is None:
            # Try without namespace
            child = element.find(tag)
        if child is None:
            # Try wildcard
            child = element.find(f"{{*}}{tag}")
        return child
