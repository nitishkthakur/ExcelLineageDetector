"""Extractor for VBA code connections."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection
from lineage.parsers.sql_parser import parse as parse_sql
from lineage.parsers.connection_string import parse as parse_cs, format_source_label


# Compile patterns once
CS_PATTERN = re.compile(
    r'(?i)(Provider|Data Source|Server|Initial Catalog|Database|DSN)\s*=\s*([^;\"\'\\n]+)',
    re.IGNORECASE,
)
SQL_PATTERN = re.compile(
    r'(?i)\b(SELECT|INSERT|UPDATE|DELETE|EXEC(?:UTE)?)\b.{5,500}',
    re.IGNORECASE | re.DOTALL,
)
FILE_PATTERN = re.compile(
    r'[A-Za-z]:\\(?:[^\"\'\\n\\/:*?<>|]+\\)*[^\"\'\\n\\/:*?<>|]*\.(?:xlsx?|xlsm|xlsb|csv|txt|mdb|accdb)',
    re.IGNORECASE,
)
UNC_PATTERN = re.compile(
    r'\\\\[^\s\"\'\\n]+',
    re.IGNORECASE,
)
URL_PATTERN = re.compile(
    r'https?://[^\s\"\'<>)]+',
    re.IGNORECASE,
)
ADODB_PATTERN = re.compile(r'(?i)ADODB\.Connection', re.IGNORECASE)
WORKBOOKS_OPEN_PATTERN = re.compile(
    r'(?i)Workbooks\.Open\s+["\']([^"\']+)["\']',
    re.IGNORECASE,
)
CONN_STR_PATTERN = re.compile(
    r'(?i)(?:ConnectionString|\.Open)\s*[=\s]\s*["\']([^"\']+)["\']',
    re.IGNORECASE,
)


class VbaExtractor(BaseExtractor):
    """Extracts data connections from VBA code."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            if "xl/vbaProject.bin" not in zip_file.namelist():
                self.log.debug("No vbaProject.bin found")
                return connections

            vba_data = zip_file.read("xl/vbaProject.bin")
            connections.extend(self._extract_from_vba(vba_data))
        except Exception as e:
            self.log.error(f"VbaExtractor failed: {e}", exc_info=True)
        return connections

    def _extract_from_vba(self, vba_data: bytes) -> list[DataConnection]:
        """Extract connections from VBA binary data using oletools."""
        results = []
        try:
            from oletools.olevba import VBA_Parser
            vba_parser = VBA_Parser("vbaProject.bin", data=vba_data)

            if not vba_parser.detect_vba_macros():
                self.log.debug("No VBA macros detected")
                return results

            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                module_name = vba_filename or stream_path or filename
                try:
                    found = self._extract_from_module(vba_code, module_name)
                    results.extend(found)
                except Exception as e:
                    self.log.warning(f"Failed to extract from VBA module {module_name}: {e}")

        except ImportError:
            self.log.warning("oletools not available, trying raw binary scan")
            results.extend(self._extract_from_raw_bytes(vba_data))
        except Exception as e:
            self.log.warning(f"VBA parsing failed: {e}, trying raw binary scan")
            results.extend(self._extract_from_raw_bytes(vba_data))

        return results

    def _extract_from_raw_bytes(self, data: bytes) -> list[DataConnection]:
        """Fallback: scan raw bytes for string patterns."""
        results = []
        try:
            # Try to decode as latin-1 (works for most binary VBA)
            text = data.decode("latin-1", errors="replace")
            results.extend(self._extract_from_module(text, "VBA:raw"))
        except Exception as e:
            self.log.debug(f"Raw byte scan failed: {e}")
        return results

    def _extract_from_module(self, code: str, module_name: str) -> list[DataConnection]:
        """Extract connections from a VBA module's source code."""
        results = []
        lines = code.splitlines()

        # Check for ADODB connections
        has_adodb = bool(ADODB_PATTERN.search(code))

        # Extract full connection strings
        for match in CONN_STR_PATTERN.finditer(code):
            cs_str = match.group(1)
            if len(cs_str) < 5:
                continue
            line_num = code[:match.start()].count('\n') + 1
            location = f"VBA:{module_name}:{line_num}"

            parsed_cs = parse_cs(cs_str)
            source = format_source_label(parsed_cs)
            sub_type = parsed_cs.get("_sub_type", "odbc")

            conn = DataConnection(
                id=DataConnection.make_id("vba", cs_str[:50], location),
                category="database",
                sub_type=sub_type,
                source=source or cs_str[:60],
                raw_connection=cs_str,
                location=location,
                metadata={
                    "module": module_name,
                    "has_adodb": has_adodb,
                    "parsed": {k: v for k, v in parsed_cs.items() if k.startswith("_")},
                },
                confidence=0.85,
            )
            results.append(conn)

        # Extract individual connection string components if no full string found
        if not results and has_adodb:
            cs_parts = {}
            for match in CS_PATTERN.finditer(code):
                key = match.group(1).lower()
                value = match.group(2).strip().strip('"\'')
                cs_parts[key] = value

            if cs_parts:
                cs_str = "; ".join(f"{k}={v}" for k, v in cs_parts.items())
                parsed_cs = parse_cs(cs_str)
                source = format_source_label(parsed_cs)
                location = f"VBA:{module_name}"

                conn = DataConnection(
                    id=DataConnection.make_id("vba", cs_str[:50], location),
                    category="database",
                    sub_type=parsed_cs.get("_sub_type", "odbc"),
                    source=source or "VBA Database Connection",
                    raw_connection=cs_str,
                    location=location,
                    metadata={"module": module_name, "components": cs_parts},
                    confidence=0.7,
                )
                results.append(conn)

        # Extract SQL statements
        seen_sql = set()
        for match in SQL_PATTERN.finditer(code):
            sql = match.group(0).strip()
            # Clean up SQL (remove VBA continuation characters)
            sql = re.sub(r'\s+&\s+_\s*\n\s*"', ' ', sql)
            sql = sql.strip('"\'').strip()
            if len(sql) < 15 or sql in seen_sql:
                continue
            seen_sql.add(sql[:50])

            line_num = code[:match.start()].count('\n') + 1
            location = f"VBA:{module_name}:{line_num}"

            parsed_query = parse_sql(sql)
            conn = DataConnection(
                id=DataConnection.make_id("vba", sql[:50], location),
                category="vba",
                sub_type="sql",
                source=f"SQL in {module_name}",
                raw_connection=sql[:200],
                location=location,
                query_text=sql,
                parsed_query=parsed_query,
                metadata={"module": module_name},
                confidence=0.8,
            )
            results.append(conn)

        # Extract file paths
        seen_files = set()
        for match in FILE_PATTERN.finditer(code):
            path = match.group(0)
            if path in seen_files:
                continue
            seen_files.add(path)

            line_num = code[:match.start()].count('\n') + 1
            location = f"VBA:{module_name}:{line_num}"

            conn = DataConnection(
                id=DataConnection.make_id("file", path, location),
                category="file",
                sub_type="local_file",
                source=path,
                raw_connection=path,
                location=location,
                metadata={"module": module_name},
                confidence=0.85,
            )
            results.append(conn)

        # Extract UNC paths
        seen_unc = set()
        for match in UNC_PATTERN.finditer(code):
            path = match.group(0).rstrip('.,;)')
            if path in seen_unc or len(path) < 5:
                continue
            seen_unc.add(path)

            line_num = code[:match.start()].count('\n') + 1
            location = f"VBA:{module_name}:{line_num}"

            conn = DataConnection(
                id=DataConnection.make_id("file", path, location),
                category="file",
                sub_type="unc_path",
                source=path,
                raw_connection=path,
                location=location,
                metadata={"module": module_name},
                confidence=0.8,
            )
            results.append(conn)

        # Extract URLs
        seen_urls = set()
        for match in URL_PATTERN.finditer(code):
            url = match.group(0).rstrip('.,;)')
            if url in seen_urls:
                continue
            seen_urls.add(url)

            line_num = code[:match.start()].count('\n') + 1
            location = f"VBA:{module_name}:{line_num}"

            conn = DataConnection(
                id=DataConnection.make_id("web", url, location),
                category="web",
                sub_type="url",
                source=url[:100],
                raw_connection=url,
                location=location,
                metadata={"module": module_name},
                confidence=0.75,
            )
            results.append(conn)

        # Extract Workbooks.Open patterns
        seen_wb = set()
        for match in WORKBOOKS_OPEN_PATTERN.finditer(code):
            path = match.group(1)
            if path in seen_wb:
                continue
            seen_wb.add(path)

            line_num = code[:match.start()].count('\n') + 1
            location = f"VBA:{module_name}:{line_num}"

            conn = DataConnection(
                id=DataConnection.make_id("file", path, location),
                category="file",
                sub_type="workbooks_open",
                source=path,
                raw_connection=path,
                location=location,
                metadata={"module": module_name},
                confidence=0.9,
            )
            results.append(conn)

        return results
