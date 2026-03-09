"""Extractor for URLs and file paths in cell comments."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

URL_PATTERN = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
FILE_PATTERN = re.compile(
    r"[A-Za-z]:\\[^\s<>\"']+|\\\\[^\s<>\"']+",
    re.IGNORECASE,
)


class CommentsExtractor(BaseExtractor):
    """Extracts URLs and file paths embedded in cell comments."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            connections.extend(self._extract_from_legacy_comments(zip_file))
        except Exception as e:
            self.log.error(f"CommentsExtractor (legacy) failed: {e}", exc_info=True)

        try:
            connections.extend(self._extract_from_threaded_comments(zip_file))
        except Exception as e:
            self.log.error(f"CommentsExtractor (threaded) failed: {e}", exc_info=True)

        return connections

    def _extract_from_legacy_comments(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract from legacy xl/comments*.xml files."""
        results = []
        comment_files = [n for n in zip_file.namelist()
                         if re.match(r"xl/comments(?:/comment)?\d*\.xml$", n)]

        for comment_file in comment_files:
            try:
                root = self._read_xml(zip_file, comment_file)
                if root is None:
                    continue

                # Get sheet name from file index
                match = re.search(r"comments(\d+)\.xml", comment_file)
                sheet_idx = match.group(1) if match else "1"
                sheet_name = self._get_sheet_name_by_index(zip_file, sheet_idx)

                # Find all text elements
                text_els = root.findall(f".//{{{NS}}}t")
                if not text_els:
                    text_els = root.findall(".//t")
                if not text_els:
                    text_els = root.findall(".//{*}t")

                # Also look for comment elements to get cell references
                comment_map = self._build_comment_ref_map(root)

                for text_el in text_els:
                    text = text_el.text or ""
                    if not text.strip():
                        continue

                    found = self._extract_connections_from_text(text, sheet_name, comment_file)
                    results.extend(found)

            except Exception as e:
                self.log.warning(f"Failed to parse {comment_file}: {e}")

        return results

    def _extract_from_threaded_comments(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract from xl/threadedComments/threadedComment*.xml files."""
        results = []
        tc_files = [n for n in zip_file.namelist()
                    if re.match(r"xl/threadedComments/threadedComment\d*\.xml$", n)]

        for tc_file in tc_files:
            try:
                root = self._read_xml(zip_file, tc_file)
                if root is None:
                    continue

                # Find text elements
                text_els = root.findall(".//{*}t")
                if not text_els:
                    text_els = root.findall(".//t")

                for text_el in text_els:
                    text = text_el.text or ""
                    if not text.strip():
                        continue

                    found = self._extract_connections_from_text(text, "comment", tc_file)
                    results.extend(found)

            except Exception as e:
                self.log.warning(f"Failed to parse threaded comment {tc_file}: {e}")

        return results

    def _extract_connections_from_text(
        self, text: str, sheet_name: str, source_file: str
    ) -> list[DataConnection]:
        """Extract URLs and file paths from comment text."""
        results = []

        # Extract URLs
        for url_match in URL_PATTERN.finditer(text):
            url = url_match.group(0).rstrip('.,;)')
            if len(url) < 10:
                continue

            location = f"{sheet_name}:comment" if sheet_name else source_file
            conn = DataConnection(
                id=DataConnection.make_id("hyperlink", url, location),
                category="hyperlink",
                sub_type="comment_url",
                source=url[:100],
                raw_connection=url,
                location=location,
                metadata={"source_file": source_file, "comment_text": text[:200]},
                confidence=0.7,
            )
            results.append(conn)

        # Extract file paths
        for file_match in FILE_PATTERN.finditer(text):
            path = file_match.group(0).rstrip('.,;)')
            if len(path) < 5:
                continue

            location = f"{sheet_name}:comment" if sheet_name else source_file
            sub_type = "unc_path" if path.startswith("\\\\") else "local_file"

            conn = DataConnection(
                id=DataConnection.make_id("file", path, location),
                category="file",
                sub_type=sub_type,
                source=path,
                raw_connection=path,
                location=location,
                metadata={"source_file": source_file, "comment_text": text[:200]},
                confidence=0.6,
            )
            results.append(conn)

        return results

    def _get_sheet_name_by_index(self, zip_file: zipfile.ZipFile, idx: str) -> str:
        """Get sheet name by its numeric index."""
        try:
            wb_root = self._read_xml(zip_file, "xl/workbook.xml")
            if wb_root is None:
                return f"Sheet{idx}"

            sheets = wb_root.findall(f".//{{{NS}}}sheet")
            if not sheets:
                sheets = wb_root.findall(".//sheet")
            if not sheets:
                sheets = wb_root.findall(".//{*}sheet")

            if sheets:
                # Comments are 1-indexed and correspond to sheet order
                idx_int = int(idx) - 1
                if 0 <= idx_int < len(sheets):
                    return sheets[idx_int].get("name", f"Sheet{idx}")
        except Exception:
            pass
        return f"Sheet{idx}"

    def _build_comment_ref_map(self, root) -> dict[str, str]:
        """Build a map of comment author/ref information."""
        comment_map = {}
        try:
            comment_list = root.find(f".//{{{NS}}}commentList")
            if comment_list is None:
                comment_list = root.find(".//commentList")
            if comment_list is None:
                comment_list = root.find(".//{*}commentList")

            if comment_list is not None:
                comment_els = comment_list.findall(f"{{{NS}}}comment")
                if not comment_els:
                    comment_els = comment_list.findall("comment")
                if not comment_els:
                    comment_els = comment_list.findall("{*}comment")

                for comment_el in comment_els:
                    ref = comment_el.get("ref", "")
                    author_id = comment_el.get("authorId", "")
                    comment_map[ref] = author_id
        except Exception:
            pass
        return comment_map
