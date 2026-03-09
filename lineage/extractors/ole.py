"""Extractor for OLE linked objects."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection


OLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


class OleExtractor(BaseExtractor):
    """Extracts OLE linked object connections."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            connections.extend(self._extract_from_drawing_rels(zip_file))
        except Exception as e:
            self.log.error(f"OleExtractor (drawings) failed: {e}", exc_info=True)

        try:
            connections.extend(self._extract_from_sheet_rels(zip_file))
        except Exception as e:
            self.log.error(f"OleExtractor (sheets) failed: {e}", exc_info=True)

        return connections

    def _extract_from_drawing_rels(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract OLE objects from drawing relationship files."""
        results = []
        names = zip_file.namelist()

        # Find drawing rels files
        drawing_rels = [n for n in names
                        if re.match(r"xl/drawings/_rels/drawing\d+\.xml\.rels$", n)]

        for rels_file in drawing_rels:
            try:
                rels = self._read_rels(zip_file, rels_file)
                for rid, rel_info in rels.items():
                    rel_type = rel_info.get("type", "")
                    if OLE_REL_TYPE not in rel_type:
                        continue

                    target = rel_info["target"]
                    target_mode = rel_info.get("mode", "Internal")

                    location = rels_file
                    conn = self._make_ole_connection(target, target_mode, location, rid)
                    if conn:
                        results.append(conn)
            except Exception as e:
                self.log.warning(f"Failed to parse {rels_file}: {e}")

        return results

    def _extract_from_sheet_rels(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract OLE objects from sheet relationship files."""
        results = []
        names = zip_file.namelist()

        # Find sheet rels files
        sheet_rels = [n for n in names
                      if re.match(r"xl/worksheets/_rels/sheet\d+\.xml\.rels$", n)]

        for rels_file in sheet_rels:
            try:
                rels = self._read_rels(zip_file, rels_file)
                for rid, rel_info in rels.items():
                    rel_type = rel_info.get("type", "")
                    if OLE_REL_TYPE not in rel_type:
                        continue

                    target = rel_info["target"]
                    target_mode = rel_info.get("mode", "Internal")

                    # Get sheet name from rels file name
                    match = re.search(r"sheet(\d+)\.xml\.rels", rels_file)
                    sheet_ref = f"Sheet{match.group(1)}" if match else "worksheet"
                    location = f"{sheet_ref}:{rels_file}"

                    conn = self._make_ole_connection(target, target_mode, location, rid)
                    if conn:
                        results.append(conn)
            except Exception as e:
                self.log.warning(f"Failed to parse {rels_file}: {e}")

        return results

    def _make_ole_connection(
        self, target: str, target_mode: str, location: str, rid: str
    ) -> DataConnection | None:
        """Create a DataConnection for an OLE object."""
        if not target:
            return None

        is_external = target_mode == "External" or self._is_external_path(target)

        # For internal OLE objects embedded in the file, we note they exist
        # For external links, they reference external files
        sub_type = "external_link" if is_external else "embedded"

        # Extract source name from path
        source = target.split("/")[-1].split("\\")[-1] if "/" in target or "\\" in target else target
        if not source:
            source = target[:60]

        conn = DataConnection(
            id=DataConnection.make_id("ole", target, location),
            category="ole",
            sub_type=sub_type,
            source=source,
            raw_connection=target,
            location=location,
            metadata={
                "rel_id": rid,
                "target_mode": target_mode,
                "is_external": is_external,
            },
            confidence=0.9 if is_external else 0.7,
        )
        return conn

    def _is_external_path(self, path: str) -> bool:
        """Check if a path is external."""
        return bool(
            path.startswith(("http://", "https://", "file://", "\\\\")) or
            re.match(r"[A-Za-z]:\\", path)
        )
