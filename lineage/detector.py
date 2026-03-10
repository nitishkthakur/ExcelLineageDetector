"""Main detector that orchestrates all extractors."""

from __future__ import annotations
import zipfile
from pathlib import Path

from lineage.extractors.connections import ConnectionsExtractor
from lineage.extractors.powerquery import PowerQueryExtractor
from lineage.extractors.formulas import FormulasExtractor
from lineage.extractors.vba import VbaExtractor
from lineage.extractors.pivot import PivotExtractor
from lineage.extractors.querytable import QueryTableExtractor
from lineage.extractors.hyperlinks import HyperlinksExtractor
from lineage.extractors.namedranges import NamedRangesExtractor
from lineage.extractors.comments import CommentsExtractor
from lineage.extractors.metadata import MetadataExtractor
from lineage.extractors.ole import OleExtractor
from lineage.extractors.externallinks import ExternalLinksExtractor
from lineage.extractors.hardcoded import HardcodedValuesExtractor
from lineage.models import DataConnection
from lineage.utils import get_logger


class ExcelLineageDetector:
    """Forensically extracts data connections from Excel files.

    Runs all registered extractors against the target Excel file and
    returns a deduplicated list of DataConnection objects.
    """

    EXTRACTORS = [
        ConnectionsExtractor,
        PowerQueryExtractor,
        FormulasExtractor,
        ExternalLinksExtractor,   # xl/externalLinks/ - resolves external workbook paths/URLs
        VbaExtractor,
        PivotExtractor,
        QueryTableExtractor,
        HyperlinksExtractor,
        NamedRangesExtractor,
        CommentsExtractor,
        MetadataExtractor,
        OleExtractor,
        HardcodedValuesExtractor,
    ]

    def __init__(self):
        self.log = get_logger("lineage.detector")

    def detect(self, path: Path) -> list[DataConnection]:
        """Run all extractors on the given Excel file.

        Args:
            path: Path to the Excel file (.xlsx, .xlsm, .xlsb).

        Returns:
            Deduplicated list of DataConnection objects found.
        """
        path = Path(path)
        connections = []

        if not path.exists():
            self.log.error(f"File not found: {path}")
            return connections

        if not path.is_file():
            self.log.error(f"Not a file: {path}")
            return connections

        self.log.info(f"Scanning: {path}")

        try:
            with zipfile.ZipFile(path) as zf:
                # Try to load workbook with openpyxl for some extractors
                wb = None
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(
                        str(path),
                        data_only=False,
                        read_only=True,
                    )
                except Exception as e:
                    self.log.warning(f"openpyxl failed to load workbook: {e}")

                # Run each extractor
                for ExtClass in self.EXTRACTORS:
                    ext = ExtClass()
                    try:
                        found = ext.extract(zf, wb)
                        self.log.debug(f"{ExtClass.__name__}: found {len(found)} connection(s)")
                        connections.extend(found)
                    except Exception as e:
                        self.log.error(f"{ExtClass.__name__} crashed: {e}", exc_info=True)

                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

        except zipfile.BadZipFile:
            self.log.error(f"File is not a valid ZIP/XLSX archive: {path}")
        except Exception as e:
            self.log.error(f"Failed to open file: {e}", exc_info=True)

        # Deduplicate by id, keeping first occurrence
        seen: dict[str, DataConnection] = {}
        for c in connections:
            if c.id not in seen:
                seen[c.id] = c

        result = list(seen.values())
        self.log.info(f"Found {len(result)} unique connection(s) (from {len(connections)} total)")
        return result
