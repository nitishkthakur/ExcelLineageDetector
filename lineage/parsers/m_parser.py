"""Power Query M code parser for extracting data source information."""

from __future__ import annotations
import re


# Each entry: (sub_type, regex_pattern)
# Patterns are tried in order; ALL matching patterns are returned by parse_all().
PATTERNS = [
    # ── Relational databases ────────────────────────────────────────────────
    ("sql_server",      r'Sql\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("azure_sql",       r'AzureSQLDatabase\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("azure_sql_dw",    r'AzureSQLDataWarehouse\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("oracle",          r'Oracle\.Database\s*\(\s*"([^"]+)"'),
    ("postgresql",      r'PostgreSQL\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("mysql",           r'MySQL\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("db2",             r'DB2\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("teradata",        r'Teradata\.Database\s*\(\s*"([^"]+)"'),
    ("sybase",          r'Sybase\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("informix",        r'Informix\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("sap_hana",        r'SapHana\.Database\s*\(\s*"([^"]+)"'),
    ("sap_bw",          r'SapBusinessWarehouse\.Cubes\s*\(\s*"([^"]+)"'),
    ("sap_bobj",        r'SapBusinessObjects\.Universe\s*\(\s*"([^"]+)"'),
    ("snowflake",       r'Snowflake\.Databases\s*\(\s*"([^"]+)"'),
    ("bigquery",        r'GoogleBigQuery\.Database\s*\(\s*"([^"]+)"'),
    ("redshift",        r'Amazon\.Redshift\s*\(\s*"([^"]+)"'),
    ("databricks",      r'Databricks\.Catalogs\s*\(\s*"([^"]+)"'),
    ("databricks",      r'Databricks\.Contents\s*\(\s*"([^"]+)"'),
    ("mongodb",         r'MongoDB\.Find\s*\(\s*"([^"]+)"'),
    ("impala",          r'Impala\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("access",          r'Access\.Database\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("odbc",            r'Odbc\.DataSource\s*\(\s*"([^"]+)"'),
    ("odbc",            r'Odbc\.Query\s*\(\s*"([^"]+)"'),
    ("oledb",           r'OleDb\.DataSource\s*\(\s*"([^"]+)"'),
    ("oledb",           r'OleDb\.Query\s*\(\s*"([^"]+)"'),

    # ── Azure cloud sources ─────────────────────────────────────────────────
    ("azure_blob",      r'AzureStorage\.Blobs\s*\(\s*"([^"]+)"'),
    ("azure_datalake",  r'AzureStorage\.DataLake\s*\(\s*"([^"]+)"'),
    ("azure_datalake",  r'AzureStorage\.DataLakeGen2\s*\(\s*"([^"]+)"'),
    ("azure_tables",    r'AzureStorage\.Tables\s*\(\s*"([^"]+)"'),
    ("azure_cosmos",    r'AzureCosmosDB\.Feed\s*\(\s*"([^"]+)"'),
    ("azure_as",        r'AzureAnalysisServices\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"'),
    ("hdfs",            r'Hdfs\.Files\s*\(\s*"([^"]+)"'),
    ("hdinsight",       r'AzureHDInsight\.Tables\s*\(\s*"([^"]+)"'),

    # ── SharePoint ──────────────────────────────────────────────────────────
    ("sharepoint",      r'SharePoint\.Files\s*\(\s*"([^"]+)"'),
    ("sharepoint",      r'SharePoint\.Tables\s*\(\s*"([^"]+)"'),
    ("sharepoint",      r'SharePoint\.List\s*\(\s*"([^"]+)"'),

    # ── CRM / SaaS ──────────────────────────────────────────────────────────
    ("salesforce",      r'Salesforce\.Data\s*\(\s*(?:"([^"]+)")?'),
    ("salesforce",      r'Salesforce\.Reports\s*\(\s*(?:"([^"]+)")?'),
    ("dynamics",        r'CommonDataService\.Database\s*\(\s*"([^"]+)"'),
    ("dynamics",        r'Dynamics365\.BusinessCentral\s*\(\s*"([^"]+)"'),
    ("exchange",        r'Exchange\.Contents\s*\(\s*"([^"]+)"'),
    ("active_directory",r'ActiveDirectory\.Domains\s*\(\s*"([^"]+)"'),

    # ── Power BI ────────────────────────────────────────────────────────────
    ("powerbi",         r'PowerBI\.Datasets\s*\(\s*(?:\[([^\]]+)\])?'),
    ("powerbi",         r'PowerBI\.Dataflows\s*\(\s*(?:\[([^\]]+)\])?'),

    # ── Web / OData ─────────────────────────────────────────────────────────
    ("web",             r'Web\.Contents\s*\(\s*"([^"]+)"'),
    ("web",             r'Web\.BrowserContents\s*\(\s*"([^"]+)"'),
    ("odata",           r'OData\.Feed\s*\(\s*"([^"]+)"'),
    ("html_table",      r'Html\.Table\s*\('),

    # ── Local / network files ────────────────────────────────────────────────
    ("file_excel",      r'Excel\.Workbook\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_excel",      r'File\.Contents\s*\(\s*"([^"]+\.xlsx?[^"]*)"'),
    ("file_csv",        r'Csv\.Document\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_json",       r'Json\.Document\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_xml",        r'Xml\.Document\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_xml",        r'Xml\.Tables\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_pdf",        r'Pdf\.Tables\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_text",       r'Lines\.FromBinary\s*\(.*?File\.Contents\s*\(\s*"([^"]+)"'),
    ("file_text",       r'File\.Contents\s*\(\s*"([^"]+)"'),
    ("folder",          r'Folder\.Files\s*\(\s*"([^"]+)"'),
    ("folder",          r'Folder\.Contents\s*\(\s*"([^"]+)"'),

    # ── Self-reference ───────────────────────────────────────────────────────
    ("self_workbook",   r'Excel\.CurrentWorkbook\s*\('),
]

# Patterns where group 1 = server, group 2 = database
_TWO_GROUP_TYPES = frozenset([
    "sql_server", "azure_sql", "azure_sql_dw", "postgresql", "mysql",
    "db2", "sybase", "informix", "impala", "azure_as",
])


def parse_all(m_code: str) -> list[dict]:
    """Parse Power Query M code and return ALL data source references found.

    Returns:
        List of dicts, each with keys: sub_type, source, details
    """
    if not m_code or not m_code.strip():
        return []

    results = []
    seen: set[str] = set()

    for sub_type, pattern in PATTERNS:
        for match in re.finditer(pattern, m_code, re.IGNORECASE | re.DOTALL):
            groups = match.groups()
            if sub_type in _TWO_GROUP_TYPES and len(groups) >= 2 and groups[1]:
                server, database = groups[0] or "", groups[1] or ""
                source = f"{server}/{database}"
                details: dict = {"server": server, "database": database}
            else:
                source = groups[0] if groups and groups[0] else ""
                details = {"path_or_url": source} if source else {}

            key = f"{sub_type}|{source}"
            if key not in seen:
                seen.add(key)
                results.append({"sub_type": sub_type, "source": source, "details": details})

    # Fallback: detect any dot-notation function call if nothing matched
    if not results:
        func_match = re.search(r'=\s*([A-Za-z]+\.[A-Za-z]+)\s*\(', m_code)
        if func_match:
            func_name = func_match.group(1)
            results.append({
                "sub_type": "m_function",
                "source": func_name,
                "details": {"function": func_name},
            })

    return results


def parse(m_code: str) -> dict:
    """Parse Power Query M code - returns first data source found.

    Kept for backward compatibility. Prefer parse_all() for full coverage.
    """
    results = parse_all(m_code)
    if results:
        return results[0]
    return {"sub_type": "unknown", "source": "", "details": {}}
