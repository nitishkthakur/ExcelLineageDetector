"""Extractor for xl/externalLinks/ - the primary storage for external workbook paths.

Excel stores external workbook references (including SharePoint, OneDrive, UNC, and
local file paths) in xl/externalLinks/externalLink*.xml and their .rels files.
When a formula uses ='[budget.xlsx]Sheet1'!A1 or =[1]Sheet1!A1, the resolved path
(which may be a full SharePoint/OneDrive URL) lives here.
"""

from __future__ import annotations
import re
import urllib.parse
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
EXTERNAL_LINK_PATH_TYPE = "externalLinkPath"


class ExternalLinksExtractor(BaseExtractor):
    """Extracts external workbook references from xl/externalLinks/.

    Covers:
    - Local file paths (C:\\path\\file.xlsx)
    - UNC paths (\\\\server\\share\\file.xlsx)
    - SharePoint URLs (https://company.sharepoint.com/...file.xlsx)
    - OneDrive URLs (https://d.docs.live.net/.../file.xlsx)
    - DDE (Dynamic Data Exchange) links
    - OLE external links
    """

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            ext_link_files = sorted([
                n for n in zip_file.namelist()
                if re.match(r"xl/externalLinks/externalLink\d+\.xml$", n)
            ])

            if not ext_link_files:
                self.log.debug("No xl/externalLinks/ entries found")
                return connections

            self.log.debug(f"Found {len(ext_link_files)} externalLink file(s)")
            for link_file in ext_link_files:
                try:
                    found = self._extract_from_link(zip_file, link_file)
                    connections.extend(found)
                except Exception as e:
                    self.log.warning(f"Failed to extract from {link_file}: {e}")

        except Exception as e:
            self.log.error(f"ExternalLinksExtractor failed: {e}", exc_info=True)

        return connections

    def _extract_from_link(
        self, zip_file: zipfile.ZipFile, link_file: str
    ) -> list[DataConnection]:
        """Extract a single externalLink entry."""
        results = []

        m = re.search(r"externalLink(\d+)\.xml$", link_file)
        link_idx = m.group(1) if m else "?"

        # The .rels file holds the actual target path or URL
        rels_path = (
            link_file
            .replace("xl/externalLinks/", "xl/externalLinks/_rels/")
            .replace(".xml", ".xml.rels")
        )
        rels = self._read_rels(zip_file, rels_path)

        target = ""
        for rel_info in rels.values():
            if EXTERNAL_LINK_PATH_TYPE in rel_info.get("type", ""):
                target = rel_info.get("target", "")
                break

        # Parse the XML body for sheet names, defined names, DDE/OLE details
        root = self._read_xml(zip_file, link_file)
        sheet_names: list[str] = []
        defined_names: list[dict] = []

        if root is not None:
            # DDE Link (Dynamic Data Exchange - e.g. link to another app like Bloomberg)
            dde_link = self._find_any(root, "ddeLink")
            if dde_link is not None:
                conn = self._make_dde_connection(dde_link, link_idx)
                if conn:
                    results.append(conn)
                return results  # DDE links have no file target

            # OLE Link (embedded/linked OLE objects)
            ole_link = self._find_any(root, "oleLink")
            if ole_link is not None:
                conn = self._make_ole_connection(ole_link, target, link_idx)
                if conn:
                    results.append(conn)
                return results

            # Standard External Book reference
            ext_book = self._find_any(root, "externalBook")
            if ext_book is not None:
                sheet_names = self._get_sheet_names(ext_book)
                defined_names = self._get_defined_names(ext_book)

        if not target:
            # No target resolved - log debug only (broken/missing link)
            self.log.debug(f"externalLink{link_idx}: no target found in rels")
            return results

        # Decode percent-encoded paths (e.g. file:///C:/My%20Files/budget.xlsx)
        target_decoded = urllib.parse.unquote(target)
        display_target = _normalize_path(target_decoded)

        category, sub_type = _classify_target(display_target)
        workbook_name = _extract_filename(display_target)

        conn = DataConnection(
            id=DataConnection.make_id(category, target, f"externalLink{link_idx}"),
            category=category,
            sub_type=sub_type,
            source=workbook_name or display_target[:80],
            raw_connection=display_target or target,
            location=f"xl/externalLinks/externalLink{link_idx}.xml",
            metadata={
                "link_index": link_idx,
                "link_file": link_file,
                "original_target": target,
                "workbook_name": workbook_name,
                "sheet_names": sheet_names,
                "defined_names": defined_names,
                # Formula index: how this link appears in cell formulas, e.g. =[1]Sheet1!A1
                "formula_index": f"[{link_idx}]",
            },
            confidence=1.0,
        )
        results.append(conn)
        return results

    # ------------------------------------------------------------------ helpers

    def _find_any(self, root, tag: str):
        """Try finding a child element with namespace, wildcard, and bare."""
        el = root.find(f"{{{NS}}}{tag}")
        if el is None:
            el = root.find(f"{{*}}{tag}")
        if el is None:
            el = root.find(tag)
        return el

    def _get_sheet_names(self, ext_book) -> list[str]:
        sheets_el = self._find_any(ext_book, "sheetNames")
        if sheets_el is None:
            return []
        for tag in [f"{{{NS}}}sheetName", "{*}sheetName", "sheetName"]:
            found = sheets_el.findall(tag)
            if found:
                return [s.get("val", "") for s in found if s.get("val")]
        return []

    def _get_defined_names(self, ext_book) -> list[dict]:
        dnames_el = self._find_any(ext_book, "definedNames")
        if dnames_el is None:
            return []
        for tag in [f"{{{NS}}}definedName", "{*}definedName", "definedName"]:
            found = dnames_el.findall(tag)
            if found:
                return [
                    {"name": dn.get("name", ""), "refersTo": dn.get("refersTo", "")}
                    for dn in found if dn.get("name")
                ]
        return []

    def _make_dde_connection(self, dde_link, link_idx: str) -> DataConnection | None:
        service = dde_link.get("ddeService", "")
        topic = dde_link.get("ddeTopic", "")
        if not service:
            return None
        raw = f"{service}|{topic}"
        return DataConnection(
            id=DataConnection.make_id("file", f"DDE:{raw}", f"externalLink{link_idx}"),
            category="file",
            sub_type="dde_link",
            source=f"DDE: {service} → {topic}" if topic else f"DDE: {service}",
            raw_connection=raw,
            location=f"xl/externalLinks/externalLink{link_idx}.xml",
            metadata={
                "link_index": link_idx,
                "dde_service": service,
                "dde_topic": topic,
                "formula_index": f"[{link_idx}]",
            },
            confidence=0.95,
        )

    def _make_ole_connection(
        self, ole_link, target: str, link_idx: str
    ) -> DataConnection | None:
        prog_id = ole_link.get("progId", "")
        raw = target or prog_id
        if not raw:
            return None
        target_decoded = _normalize_path(urllib.parse.unquote(target)) if target else ""
        return DataConnection(
            id=DataConnection.make_id("ole", raw, f"externalLink{link_idx}"),
            category="ole",
            sub_type="ole_link",
            source=f"OLE: {prog_id}" if prog_id else "OLE External Link",
            raw_connection=target_decoded or raw,
            location=f"xl/externalLinks/externalLink{link_idx}.xml",
            metadata={
                "link_index": link_idx,
                "prog_id": prog_id,
                "original_target": target,
                "formula_index": f"[{link_idx}]",
            },
            confidence=0.95,
        )


# ------------------------------------------------------------------ utilities

def _normalize_path(target: str) -> str:
    """Convert file:// URLs and other URI forms to display paths."""
    if target.startswith("file:///"):
        rest = target[8:]
        # Windows absolute path: file:///C:/... → C:\...
        if re.match(r"[A-Za-z]:/", rest):
            return rest.replace("/", "\\")
        # Bare UNC remainder (rare): treat as-is
        return rest
    if target.startswith("file://"):
        # UNC form: file://server/share/path → \\server\share\path
        return "\\\\" + target[7:].replace("/", "\\")
    return target


def _classify_target(path: str) -> tuple[str, str]:
    """Return (category, sub_type) based on the target path format."""
    pl = path.lower()
    if pl.startswith("https://") or pl.startswith("http://"):
        if "sharepoint.com" in pl or "/sharepoint/" in pl or "/_layouts/" in pl:
            return "file", "sharepoint_workbook"
        if "onedrive" in pl or "live.net" in pl or "1drv.ms" in pl:
            return "file", "onedrive_workbook"
        return "web", "http_workbook"
    if path.startswith("\\\\"):
        return "file", "unc_path"
    return "file", "external_workbook"


def _extract_filename(path: str) -> str:
    """Return the last path component (filename)."""
    if not path:
        return ""
    return path.replace("\\", "/").rstrip("/").split("/")[-1]
