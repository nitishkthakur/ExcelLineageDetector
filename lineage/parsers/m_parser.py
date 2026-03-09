"""Power Query M code parser for extracting data source information."""

from __future__ import annotations
import re


PATTERNS = [
    ("sql_server", r'Sql\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("oracle", r'Oracle\.Database\s*\(\s*"([^"]+)"'),
    ("file_excel", r'Excel\.Workbook\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_excel", r'File\.Contents\s*\(\s*"([^"]+\.xlsx?[^"]*)"'),
    ("file_csv", r'Csv\.Document\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("web", r'Web\.Contents\s*\(\s*"([^"]+)"'),
    ("odata", r'OData\.Feed\s*\(\s*"([^"]+)"'),
    ("sharepoint", r'SharePoint\.Files\s*\(\s*"([^"]+)"'),
    ("sharepoint", r'SharePoint\.Tables\s*\(\s*"([^"]+)"'),
    ("azure_blob", r'AzureStorage\.Blobs\s*\(\s*"([^"]+)"'),
    ("azure_sql", r'AzureStorage\.DataLake\s*\(\s*"([^"]+)"'),
    ("folder", r'Folder\.Files\s*\(\s*"([^"]+)"'),
    ("access", r'Access\.Database\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("postgresql", r'PostgreSQL\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("mysql", r'MySQL\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("odbc", r'Odbc\.DataSource\s*\(\s*"([^"]+)"'),
    ("oledb", r'OleDb\.DataSource\s*\(\s*"([^"]+)"'),
    ("json", r'Json\.Document\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("xml", r'Xml\.Document\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("text_file", r'File\.Contents\s*\(\s*"([^"]+)"'),
]


def parse(m_code: str) -> dict:
    """Parse Power Query M code and extract data source information.

    Returns:
        dict with keys: sub_type, source, details
    """
    if not m_code or not m_code.strip():
        return {"sub_type": "unknown", "source": "", "details": {}}

    for sub_type, pattern in PATTERNS:
        match = re.search(pattern, m_code, re.IGNORECASE | re.DOTALL)
        if match:
            groups = match.groups()
            if sub_type == "sql_server" and len(groups) >= 2:
                server, database = groups[0], groups[1]
                source = f"{server}/{database}"
                return {
                    "sub_type": sub_type,
                    "source": source,
                    "details": {"server": server, "database": database},
                }
            elif sub_type in ("postgresql", "mysql") and len(groups) >= 2:
                server, database = groups[0], groups[1]
                source = f"{server}/{database}"
                return {
                    "sub_type": sub_type,
                    "source": source,
                    "details": {"server": server, "database": database},
                }
            else:
                source = groups[0] if groups else ""
                return {
                    "sub_type": sub_type,
                    "source": source,
                    "details": {"path_or_url": source},
                }

    # Try to find any let...in block and extract source step
    let_match = re.search(r'let\s+(.*?)\s+in\s+', m_code, re.IGNORECASE | re.DOTALL)
    if let_match:
        # Check for known function calls
        func_match = re.search(r'=\s*([A-Za-z]+\.[A-Za-z]+)\s*\(', m_code)
        if func_match:
            func_name = func_match.group(1)
            return {
                "sub_type": "m_function",
                "source": func_name,
                "details": {"function": func_name},
            }

    return {"sub_type": "unknown", "source": "", "details": {}}
