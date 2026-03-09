"""Extractor for document metadata properties."""

from __future__ import annotations
import re
import zipfile

from lineage.extractors.base import BaseExtractor
from lineage.models import DataConnection


# Core properties namespace
CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"

# App properties namespace
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

# Custom properties namespace
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

URL_PATTERN = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
FILE_PATTERN = re.compile(
    r"[A-Za-z]:\\[^\s<>\"']+|\\\\[^\s<>\"']+",
    re.IGNORECASE,
)


class MetadataExtractor(BaseExtractor):
    """Extracts data connections from document metadata properties."""

    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        connections = []
        try:
            connections.extend(self._extract_core_props(zip_file))
        except Exception as e:
            self.log.error(f"MetadataExtractor (core) failed: {e}", exc_info=True)

        try:
            connections.extend(self._extract_app_props(zip_file))
        except Exception as e:
            self.log.error(f"MetadataExtractor (app) failed: {e}", exc_info=True)

        try:
            connections.extend(self._extract_custom_props(zip_file))
        except Exception as e:
            self.log.error(f"MetadataExtractor (custom) failed: {e}", exc_info=True)

        return connections

    def _extract_core_props(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract from docProps/core.xml."""
        results = []
        root = self._read_xml(zip_file, "docProps/core.xml")
        if root is None:
            return results

        props = {}

        # Creator
        creator_el = root.find(f"{{{DC_NS}}}creator")
        if creator_el is not None and creator_el.text:
            props["creator"] = creator_el.text

        # Last modified by
        last_mod_el = root.find(f"{{{CORE_NS}}}lastModifiedBy")
        if last_mod_el is not None and last_mod_el.text:
            props["lastModifiedBy"] = last_mod_el.text

        # Created date
        created_el = root.find(f"{{{DCTERMS_NS}}}created")
        if created_el is not None and created_el.text:
            props["created"] = created_el.text

        # Modified date
        modified_el = root.find(f"{{{DCTERMS_NS}}}modified")
        if modified_el is not None and modified_el.text:
            props["modified"] = modified_el.text

        if props:
            conn = DataConnection(
                id=DataConnection.make_id("metadata", "core_properties", "docProps/core.xml"),
                category="metadata",
                sub_type="core_properties",
                source="Document Core Properties",
                raw_connection="docProps/core.xml",
                location="docProps/core.xml",
                author=props.get("creator"),
                created_at=props.get("created"),
                modified_at=props.get("modified"),
                metadata=props,
                confidence=1.0,
            )
            results.append(conn)

        return results

    def _extract_app_props(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract from docProps/app.xml."""
        results = []
        root = self._read_xml(zip_file, "docProps/app.xml")
        if root is None:
            return results

        props = {}
        interesting_tags = [
            "Application", "Manager", "Company",
            "HyperlinkBase", "AppVersion", "DocSecurity",
        ]

        for tag in interesting_tags:
            el = root.find(f"{{{APP_NS}}}{tag}")
            if el is None:
                el = root.find(tag)
            if el is not None and el.text:
                props[tag] = el.text

        if not props:
            return results

        # Check HyperlinkBase for interesting value
        hyperlink_base = props.get("HyperlinkBase", "")
        if hyperlink_base and (URL_PATTERN.search(hyperlink_base) or FILE_PATTERN.search(hyperlink_base)):
            conn = DataConnection(
                id=DataConnection.make_id("metadata", hyperlink_base, "docProps/app.xml"),
                category="metadata",
                sub_type="hyperlink_base",
                source=hyperlink_base,
                raw_connection=hyperlink_base,
                location="docProps/app.xml",
                metadata=props,
                confidence=0.9,
            )
            results.append(conn)

        # General metadata connection
        conn = DataConnection(
            id=DataConnection.make_id("metadata", "app_properties", "docProps/app.xml"),
            category="metadata",
            sub_type="app_properties",
            source=props.get("Application", "Unknown Application"),
            raw_connection="docProps/app.xml",
            location="docProps/app.xml",
            metadata=props,
            confidence=1.0,
        )
        results.append(conn)

        return results

    def _extract_custom_props(self, zip_file: zipfile.ZipFile) -> list[DataConnection]:
        """Extract from docProps/custom.xml."""
        results = []
        root = self._read_xml(zip_file, "docProps/custom.xml")
        if root is None:
            return results

        # Parse custom properties
        prop_els = root.findall(f"{{{CUSTOM_NS}}}property")
        if not prop_els:
            prop_els = root.findall("property")
        if not prop_els:
            prop_els = root.findall("{*}property")

        for prop_el in prop_els:
            try:
                prop_name = prop_el.get("name", "")
                prop_fmtid = prop_el.get("fmtid", "")

                # Get value from various vt: elements
                value = None
                for vt_tag in ["lpwstr", "lpstr", "bstr"]:
                    vt_el = prop_el.find(f"{{{VT_NS}}}{vt_tag}")
                    if vt_el is None:
                        vt_el = prop_el.find(vt_tag)
                    if vt_el is not None and vt_el.text:
                        value = vt_el.text
                        break

                if not value:
                    continue

                # Check if value contains URL or file path
                has_url = bool(URL_PATTERN.search(value))
                has_path = bool(FILE_PATTERN.search(value))

                if not has_url and not has_path:
                    continue

                location = f"docProps/custom.xml:{prop_name}"
                sub_type = "url" if has_url else "file_path"
                category = "web" if has_url else "file"

                # Extract the actual URL or path
                if has_url:
                    raw_conn = URL_PATTERN.search(value).group(0)
                else:
                    raw_conn = FILE_PATTERN.search(value).group(0)

                conn = DataConnection(
                    id=DataConnection.make_id("metadata", raw_conn, location),
                    category="metadata",
                    sub_type=f"custom_property_{sub_type}",
                    source=f"{prop_name}: {value[:60]}",
                    raw_connection=raw_conn,
                    location=location,
                    metadata={
                        "property_name": prop_name,
                        "property_value": value,
                        "fmtid": prop_fmtid,
                    },
                    confidence=0.8,
                )
                results.append(conn)

            except Exception as e:
                self.log.debug(f"Error processing custom property: {e}")

        return results
