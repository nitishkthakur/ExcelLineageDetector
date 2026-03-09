"""Abstract base extractor for Excel Lineage Detector."""

from __future__ import annotations
import zipfile
from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

from lineage.models import DataConnection
from lineage.utils import get_logger

if TYPE_CHECKING:
    import openpyxl


class BaseExtractor(ABC):
    """Abstract base class for all extractors.

    Each extractor must:
    - Catch ALL exceptions internally and log them.
    - Return whatever partial results were collected.
    - Never re-raise exceptions.
    """

    def __init__(self):
        self.log = get_logger(f"lineage.{self.__class__.__name__}")

    @abstractmethod
    def extract(self, zip_file: zipfile.ZipFile, workbook) -> list[DataConnection]:
        """Extract data connections from the Excel file.

        Args:
            zip_file: The open ZipFile object for direct XML access.
            workbook: The openpyxl workbook (may be None if loading failed).

        Returns:
            List of DataConnection objects found. May be partial if errors occurred.
        """
        ...

    def _has_file(self, zip_file: zipfile.ZipFile, path: str) -> bool:
        """Check if a file exists inside the zip archive."""
        return path in zip_file.namelist()

    def _read_xml(self, zip_file: zipfile.ZipFile, path: str):
        """Read and parse an XML file from the zip archive using lxml.

        Returns lxml Element or None if not found or parse error.
        """
        try:
            from lxml import etree
            data = zip_file.read(path)
            return etree.fromstring(data)
        except KeyError:
            self.log.debug(f"File not found in zip: {path}")
            return None
        except Exception as e:
            self.log.warning(f"Failed to parse XML {path}: {e}")
            return None

    def _read_rels(self, zip_file: zipfile.ZipFile, path: str) -> dict[str, str]:
        """Read a .rels file and return a mapping of rId -> target."""
        result = {}
        try:
            from lxml import etree
            data = zip_file.read(path)
            root = etree.fromstring(data)
            ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            for rel in root.findall(f"{{{ns}}}Relationship"):
                rid = rel.get("Id", "")
                target = rel.get("Target", "")
                target_mode = rel.get("TargetMode", "Internal")
                result[rid] = {"target": target, "type": rel.get("Type", ""), "mode": target_mode}
        except KeyError:
            self.log.debug(f"Rels file not found: {path}")
        except Exception as e:
            self.log.warning(f"Failed to parse rels {path}: {e}")
        return result
