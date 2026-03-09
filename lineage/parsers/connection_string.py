"""Parser for ODBC/OLE DB connection strings."""

from __future__ import annotations
import re


# Known provider mappings
PROVIDER_MAP = {
    "sqloledb": "sql_server",
    "sqlncli": "sql_server",
    "sqlncli10": "sql_server",
    "sqlncli11": "sql_server",
    "msoledbsql": "sql_server",
    "microsoft.ace.oledb.12.0": "access",
    "microsoft.ace.oledb.16.0": "access",
    "microsoft.jet.oledb.4.0": "access",
    "oraoledb.oracle": "oracle",
    "msdaora": "oracle",
    "msdasql": "odbc",
    "ibmda400": "as400",
    "aseoledb": "sybase",
    "sybase.asoledb": "sybase",
    "mysql.data.mysqlclient": "mysql",
    "mysqlprov": "mysql",
    "postgresql": "postgresql",
    "npgsql": "postgresql",
}

# Driver mappings for ODBC DSN-less
DRIVER_MAP = {
    "sql server": "sql_server",
    "sql server native client": "sql_server",
    "odbc driver": "sql_server",
    "oracle": "oracle",
    "mysql odbc": "mysql",
    "postgresql": "postgresql",
    "sqlite3 odbc": "sqlite",
    "microsoft access driver": "access",
    "microsoft excel driver": "excel",
    "microsoft text driver": "text",
    "sybase": "sybase",
    "db2": "db2",
    "informix": "informix",
    "teradata": "teradata",
}


def parse(connection_string: str) -> dict:
    """Parse an ODBC or OLE DB connection string.

    Returns:
        dict with extracted key-value pairs plus:
        - _sub_type: detected database type
        - _server: server/host
        - _database: database name
        - _provider: provider name
        - _dsn: DSN name
    """
    if not connection_string or not connection_string.strip():
        return {}

    result = {}

    # Split by semicolons (but not those inside quotes)
    parts = _split_connection_string(connection_string)

    for part in parts:
        if '=' not in part:
            continue
        key, _, value = part.partition('=')
        key = key.strip().lower()
        value = value.strip().strip('"').strip("'")
        if key and value:
            result[key] = value

    # Normalize common keys
    normalized = {}

    # Server / Data Source
    server = (result.get("server") or result.get("data source") or
               result.get("datasource") or result.get("host") or
               result.get("address") or result.get("addr") or "")
    if server:
        normalized["_server"] = server

    # Database
    database = (result.get("database") or result.get("initial catalog") or
                 result.get("dbq") or result.get("db") or "")
    if database:
        normalized["_database"] = database

    # Provider
    provider = result.get("provider", "").lower()
    if provider:
        normalized["_provider"] = provider

    # DSN
    dsn = result.get("dsn", "")
    if dsn:
        normalized["_dsn"] = dsn

    # Driver
    driver = result.get("driver", "").lower().strip("{}")
    if driver:
        normalized["_driver"] = driver

    # Determine sub_type
    sub_type = "odbc"  # default
    if provider:
        for prov_key, prov_type in PROVIDER_MAP.items():
            if prov_key in provider:
                sub_type = prov_type
                break
    elif driver:
        for drv_key, drv_type in DRIVER_MAP.items():
            if drv_key in driver:
                sub_type = drv_type
                break
    elif dsn:
        sub_type = "odbc_dsn"

    normalized["_sub_type"] = sub_type

    # Merge original keys (lowercase)
    for k, v in result.items():
        if k not in normalized:
            normalized[k] = v

    return normalized


def _split_connection_string(cs: str) -> list[str]:
    """Split connection string by semicolons, respecting quoted values."""
    parts = []
    current = []
    in_quote = None

    for char in cs:
        if char in ('"', "'") and in_quote is None:
            in_quote = char
            current.append(char)
        elif char == in_quote:
            in_quote = None
            current.append(char)
        elif char == ';' and in_quote is None:
            parts.append(''.join(current).strip())
            current = []
        else:
            current.append(char)

    if current:
        parts.append(''.join(current).strip())

    return [p for p in parts if p]


def format_source_label(parsed: dict) -> str:
    """Format a human-readable source label from parsed connection string."""
    server = parsed.get("_server", "")
    database = parsed.get("_database", "")
    dsn = parsed.get("_dsn", "")
    sub_type = parsed.get("_sub_type", "")

    if server and database:
        return f"{server}/{database}"
    elif server:
        return server
    elif dsn:
        return f"DSN:{dsn}"
    elif database:
        return database
    elif sub_type:
        return sub_type
    return "unknown"
