"""Extractor for Pivot Table data sources."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class PivotExtractor(BaseExtractor):
    """Extracts data source connections from Pivot Tables."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            # Find all pivot table files
            pivot_files = [n for n in zip_file.namelist()
                           if re.match(r"xl/pivotTables/pivotTable\d+\.xml$", n)]

            if not pivot_files:
                self.log.debug("No pivot table files found")
                return connections

            # Build cache definition map
            cache_defs = self._load_cache_definitions(zip_file)

            for pivot_file in pivot_files:
                try:
                    found = self._extract_from_pivot(zip_file, pivot_file, cache_defs)
                    connections.extend(found)
                except Exception as e:
                    self.log.warning(f"Failed to extract from {pivot_file}: {e}")

        except Exception as e:
            self.log.error(f"PivotExtractor failed: {e}", exc_info=True)

        return connections

    def _load_cache_definitions(self, zip_file: zipfile.ZipFile) -> dict[str, dict]:
        """Load all pivot cache definitions indexed by their file number."""
        cache_defs = {}
        cache_files = [n for n in zip_file.namelist()
                       if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", n)]

        for cache_file in cache_files:
            try:
                root = self._read_xml(zip_file, cache_file)
                if root is None:
                    continue

                # Extract cache ID from filename
                match = re.search(r"pivotCacheDefinition(\d+)\.xml", cache_file)
                cache_num = match.group(1) if match else "0"

                cache_info = self._parse_cache_definition(root, cache_file, cache_num)
                if cache_info:
                    cache_defs[cache_num] = cache_info

            except Exception as e:
                self.log.warning(f"Failed to load cache definition {cache_file}: {e}")

        return cache_defs

    def _parse_cache_definition(self, root, cache_file: str, cache_num: str) -> dict | None:
        """Parse a pivot cache definition XML."""
        # Look for cacheSource element
        cache_source = root.find(f"{{{NS}}}cacheSource")
        if cache_source is None:
            cache_source = root.find("cacheSource")
        if cache_source is None:
            cache_source = root.find("{*}cacheSource")

        if cache_source is None:
            return None

        source_type = cache_source.get("type", "worksheet")

        info = {
            "cache_num": cache_num,
            "cache_file": cache_file,
            "source_type": source_type,
        }

        if source_type == "worksheet":
            ws_source = cache_source.find(f"{{{NS}}}worksheetSource")
            if ws_source is None:
                ws_source = cache_source.find("worksheetSource")
            if ws_source is None:
                ws_source = cache_source.find("{*}worksheetSource")

            if ws_source is not None:
                info["worksheet"] = ws_source.get("sheet", "")
                info["ref"] = ws_source.get("ref", "")
                info["rid"] = ws_source.get(f"{{{REL_NS}}}id") or ws_source.get("r:id", "")

        elif source_type == "external":
            db_pr = cache_source.find(f"{{{NS}}}dbPr")
            if db_pr is None:
                db_pr = cache_source.find("dbPr")
            if db_pr is None:
                db_pr = cache_source.find("{*}dbPr")

            if db_pr is not None:
                info["connection"] = db_pr.get("connection", "")
                info["command"] = db_pr.get("command", "")

        elif source_type == "consolidation":
            info["is_consolidation"] = True

        return info

    def _extract_from_pivot(
        self,
        zip_file: zipfile.ZipFile,
        pivot_file: str,
        cache_defs: dict,
    ) -> list[DataConnection]:
        """Extract connections from a single pivot table file."""
        results = []

        root = self._read_xml(zip_file, pivot_file)
        if root is None:
            return results

        # Get pivot table attributes
        pivot_name = root.get("name", "PivotTable")
        cache_id = root.get("cacheId", "")

        # Get sheet name from pivot file relationships
        sheet_name = self._get_pivot_sheet_name(zip_file, pivot_file)
        location = f"{sheet_name}:PivotTable:{pivot_name}" if sheet_name else f"PivotTable:{pivot_name}"

        # Find matching cache definition
        cache_info = None
        if cache_id:
            # Map cache_id to cache definition number
            # The cacheId in pivot table matches the cacheId attribute in workbook.xml pivotCaches
            cache_info = self._find_cache_by_id(zip_file, cache_id, cache_defs)

        if not cache_info:
            # Try to find any cache info
            for cd in cache_defs.values():
                cache_info = cd
                break

        if not cache_info:
            # Create a minimal connection indicating pivot exists
            conn = DataConnection(
                id=DataConnection.make_id("pivot", pivot_name, location),
                category="pivot",
                sub_type="internal",
                source=pivot_name,
                raw_connection=pivot_name,
                location=location,
                metadata={"cache_id": cache_id, "pivot_name": pivot_name},
                confidence=0.6,
            )
            results.append(conn)
            return results

        source_type = cache_info.get("source_type", "worksheet")

        if source_type == "worksheet":
            worksheet = cache_info.get("worksheet", "")
            ref = cache_info.get("ref", "")
            source = f"{worksheet}!{ref}" if worksheet and ref else worksheet or "internal"

            conn = DataConnection(
                id=DataConnection.make_id("pivot", source, location),
                category="pivot",
                sub_type="worksheet",
                source=source,
                raw_connection=source,
                location=location,
                metadata={
                    "pivot_name": pivot_name,
                    "cache_id": cache_id,
                    "source_sheet": worksheet,
                    "source_range": ref,
                },
            )
            results.append(conn)

        elif source_type == "external":
            connection = cache_info.get("connection", "")
            command = cache_info.get("command", "")
            source = connection or "external"

            conn = DataConnection(
                id=DataConnection.make_id("pivot", connection, location),
                category="pivot",
                sub_type="external_db",
                source=source,
                raw_connection=connection,
                location=location,
                query_text=command or None,
                metadata={
                    "pivot_name": pivot_name,
                    "cache_id": cache_id,
                },
            )
            results.append(conn)

        elif cache_info.get("is_consolidation"):
            conn = DataConnection(
                id=DataConnection.make_id("pivot", pivot_name + "_consolidation", location),
                category="pivot",
                sub_type="consolidation",
                source=f"{pivot_name} (consolidation)",
                raw_connection=pivot_name,
                location=location,
                metadata={"pivot_name": pivot_name, "cache_id": cache_id},
                confidence=0.7,
            )
            results.append(conn)

        return results

    def _get_pivot_sheet_name(self, zip_file: zipfile.ZipFile, pivot_file: str) -> str:
        """Try to determine which sheet contains this pivot table."""
        try:
            # Find the pivot table's relationship file
            pivot_rels_file = pivot_file.replace(
                "xl/pivotTables/", "xl/pivotTables/_rels/"
            ).replace(".xml", ".xml.rels")

            if pivot_rels_file in zip_file.namelist():
                rels = self._read_rels(zip_file, pivot_rels_file)
                for rid, rel_info in rels.items():
                    target = rel_info["target"]
                    if "worksheets" in target:
                        # Extract sheet number
                        match = re.search(r"sheet(\d+)\.xml", target)
                        if match:
                            sheet_num = match.group(1)
                            return f"Sheet{sheet_num}"
        except Exception:
            pass
        return ""

    def _find_cache_by_id(
        self, zip_file: zipfile.ZipFile, cache_id: str, cache_defs: dict
    ) -> dict | None:
        """Find a cache definition by its cacheId from workbook.xml."""
        try:
            wb_root = self._read_xml(zip_file, "xl/workbook.xml")
            if wb_root is None:
                return None

            # Find pivotCaches element
            pivot_caches = wb_root.find(f"{{{NS}}}pivotCaches")
            if pivot_caches is None:
                pivot_caches = wb_root.find("pivotCaches")
            if pivot_caches is None:
                pivot_caches = wb_root.find("{*}pivotCaches")

            if pivot_caches is None:
                return None

            cache_els = pivot_caches.findall(f"{{{NS}}}pivotCache")
            if not cache_els:
                cache_els = pivot_caches.findall("pivotCache")
            if not cache_els:
                cache_els = pivot_caches.findall("{*}pivotCache")

            for cache_el in cache_els:
                if cache_el.get("cacheId") == cache_id:
                    rid = (cache_el.get(f"{{{REL_NS}}}id") or
                           cache_el.get("r:id") or "")
                    # Find the cache definition file number
                    for cache_num, cache_info in cache_defs.items():
                        cache_file = cache_info.get("cache_file", "")
                        if f"pivotCacheDefinition{cache_num}" in cache_file:
                            return cache_info

        except Exception as e:
            self.log.debug(f"Failed to find cache by id {cache_id}: {e}")

        return None
