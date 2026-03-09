"""Extractor for hyperlinks in worksheets."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


class HyperlinksExtractor(BaseExtractor):
    """Extracts external hyperlinks from worksheets."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            sheet_map = self._get_sheet_map(zip_file)
            for sheet_name, sheet_file in sheet_map.items():
                try:
                    found = self._extract_from_sheet(zip_file, sheet_name, sheet_file)
                    connections.extend(found)
                except Exception as e:
                    self.log.warning(f"Failed to extract hyperlinks from {sheet_name}: {e}")
        except Exception as e:
            self.log.error(f"HyperlinksExtractor failed: {e}", exc_info=True)
        return connections

    def _get_sheet_map(self, zip_file: zipfile.ZipFile) -> dict[str, str]:
        """Get mapping of sheet name -> file path."""
        sheet_map = {}
        try:
            wb_root = self._read_xml(zip_file, "xl/workbook.xml")
            if wb_root is None:
                return sheet_map

            sheets = wb_root.findall(f"{{{NS}}}sheets/{{{NS}}}sheet")
            if not sheets:
                sheets = wb_root.findall(".//sheets/sheet")
            if not sheets:
                sheets = wb_root.findall(".//{*}sheet")

            rels = self._read_rels(zip_file, "xl/_rels/workbook.xml.rels")

            for sheet_el in sheets:
                name = sheet_el.get("name", "")
                rid = (sheet_el.get(f"{{{REL_NS}}}id") or
                       sheet_el.get("r:id") or sheet_el.get("id", ""))

                if rid in rels:
                    target = rels[rid]["target"]
                    # Strip leading slash from absolute paths (e.g. /xl/worksheets/sheet1.xml)
                    target = target.lstrip("/")
                    if not target.startswith("xl/"):
                        target = f"xl/{target}"
                    sheet_map[name] = target
        except Exception as e:
            self.log.warning(f"Failed to get sheet map: {e}")

        # Fallback
        if not sheet_map:
            for name in zip_file.namelist():
                if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                    m = re.search(r"sheet(\d+)\.xml", name)
                    if m:
                        sheet_map[f"Sheet{m.group(1)}"] = name

        return sheet_map

    def _extract_from_sheet(
        self, zip_file: zipfile.ZipFile, sheet_name: str, sheet_file: str
    ) -> list[DataConnection]:
        """Extract hyperlinks from a single sheet."""
        results = []

        if sheet_file not in zip_file.namelist():
            return results

        root = self._read_xml(zip_file, sheet_file)
        if root is None:
            return results

        # Load relationships for this sheet
        # Rels file is at xl/worksheets/_rels/sheetN.xml.rels
        sheet_filename = sheet_file.split("/")[-1]
        rels_file = f"xl/worksheets/_rels/{sheet_filename}.rels"
        rels = self._read_rels(zip_file, rels_file)

        # Find hyperlink elements
        hyperlinks = root.findall(f".//{{{NS}}}hyperlink")
        if not hyperlinks:
            hyperlinks = root.findall(".//hyperlink")
        if not hyperlinks:
            hyperlinks = root.findall(".//{*}hyperlink")

        for hl in hyperlinks:
            try:
                ref = hl.get("ref", "")
                rid = (hl.get(f"{{{REL_NS}}}id") or
                       hl.get("r:id") or hl.get("id") or "")
                display = hl.get("display", "") or hl.get("tooltip", "")
                location_attr = hl.get("location", "")  # internal anchor

                # Get the actual URL/target from rels
                url = ""
                is_external = False

                if rid and rid in rels:
                    rel_info = rels[rid]
                    url = rel_info["target"]
                    is_external = rel_info.get("mode", "") == "External"

                # Skip internal hyperlinks (no URL or just anchor)
                if not url or (not is_external and not self._is_external_url(url)):
                    continue

                location = f"{sheet_name}:{ref}" if ref else sheet_name

                sub_type = self._get_hyperlink_subtype(url)
                category = "web" if url.startswith(("http://", "https://")) else "file"

                conn = DataConnection(
                    id=DataConnection.make_id("hyperlink", url, location),
                    category="hyperlink",
                    sub_type=sub_type,
                    source=display or url[:80],
                    raw_connection=url,
                    location=location,
                    metadata={
                        "cell_ref": ref,
                        "display_text": display,
                        "sheet": sheet_name,
                        "is_external": is_external,
                    },
                )
                results.append(conn)

            except Exception as e:
                self.log.debug(f"Error processing hyperlink: {e}")

        return results

    def _is_external_url(self, url: str) -> bool:
        """Check if URL is external (http, file, UNC, mailto)."""
        return bool(
            url.startswith(("http://", "https://", "file://", "\\\\", "mailto:")) or
            re.match(r"[A-Za-z]:\\", url)
        )

    def _get_hyperlink_subtype(self, url: str) -> str:
        """Determine hyperlink sub-type from URL."""
        if url.startswith(("http://", "https://")):
            return "http"
        elif url.startswith("file://"):
            return "file_url"
        elif url.startswith("\\\\"):
            return "unc_path"
        elif url.startswith("mailto:"):
            return "mailto"
        elif re.match(r"[A-Za-z]:\\", url):
            return "local_file"
        return "hyperlink"
