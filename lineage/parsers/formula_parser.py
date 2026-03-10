"""Parser for Excel external formula references."""

from __future__ import annotations
import re


# Pattern to match external workbook references like:
# '[workbook.xlsx]Sheet1'!A1
# 'C:\path\[workbook.xlsx]Sheet1'!A1
# '\\server\share\[workbook.xlsx]Sheet1'!A1
# 'https://company.sharepoint.com/sites/dept/Docs/[workbook.xlsx]Sheet1'!A1
EXTERNAL_REF_PATTERN = re.compile(
    r"'?([A-Za-z]:\\[^'[]*|\\\\[^'[]*|https?://[^'[]+)?(?:\[([^\]]+\.(?:xlsx?|xlsm|xlsb|csv))\])([^'!]*)(?:')?!([A-Z$][A-Z0-9$:]*)",
    re.IGNORECASE,
)

# Simpler pattern for just [workbook.xlsx] references
SIMPLE_WORKBOOK_PATTERN = re.compile(
    r"\[([^\]]+\.(?:xlsx?|xlsm|xlsb|csv))\]",
    re.IGNORECASE,
)

# UNC path pattern
UNC_PATTERN = re.compile(r"'(\\\\[^']+)'", re.IGNORECASE)

# Local drive path pattern
LOCAL_PATH_PATTERN = re.compile(r"'([A-Za-z]:\\[^']+)'", re.IGNORECASE)


def parse(formula: str) -> dict | None:
    """Parse an Excel formula external reference string.

    Returns:
        dict with keys: workbook_path, workbook_name, sheet, cell_ref
        or None if no external reference found.
    """
    if not formula:
        return None

    # Try full match first
    match = EXTERNAL_REF_PATTERN.search(formula)
    if match:
        path, workbook, sheet, cell = match.groups()
        return {
            "workbook_path": (path or "").strip("'"),
            "workbook_name": workbook or "",
            "sheet": sheet.strip("'") if sheet else "",
            "cell_ref": cell or "",
            "full_ref": match.group(0),
        }

    # Try simple workbook pattern
    wb_match = SIMPLE_WORKBOOK_PATTERN.search(formula)
    if wb_match:
        workbook = wb_match.group(1)
        # Try to extract sheet name after the bracket
        rest = formula[wb_match.end():]
        sheet_match = re.match(r"([^'!]+)'?!", rest)
        sheet = sheet_match.group(1).strip("'") if sheet_match else ""
        return {
            "workbook_path": "",
            "workbook_name": workbook,
            "sheet": sheet,
            "cell_ref": "",
            "full_ref": wb_match.group(0),
        }

    # Check for UNC path
    unc_match = UNC_PATTERN.search(formula)
    if unc_match:
        path = unc_match.group(1)
        return {
            "workbook_path": path,
            "workbook_name": path.split("\\")[-1] if "\\" in path else path,
            "sheet": "",
            "cell_ref": "",
            "full_ref": unc_match.group(0),
        }

    # Check for local path
    local_match = LOCAL_PATH_PATTERN.search(formula)
    if local_match:
        path = local_match.group(1)
        return {
            "workbook_path": path,
            "workbook_name": path.split("\\")[-1] if "\\" in path else path,
            "sheet": "",
            "cell_ref": "",
            "full_ref": local_match.group(0),
        }

    return None
