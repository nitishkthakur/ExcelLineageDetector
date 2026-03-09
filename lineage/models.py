from __future__ import annotations
import hashlib
import json
from dataclasses import dataclass, field, asdict
from typing import Any


@dataclass
class ParsedQuery:
    tables: list[str] = field(default_factory=list)
    columns: list[str] = field(default_factory=list)
    joins: list[dict] = field(default_factory=list)
    filters: list[str] = field(default_factory=list)
    raw_sql: str = ""


@dataclass
class DataConnection:
    id: str
    category: str          # database|file|web|powerquery|vba|pivot|formula|hyperlink|ole|metadata
    sub_type: str          # odbc|oledb|sql_server|oracle|xlsx|csv|odata|webservice|rtd|...
    source: str            # human-readable label
    raw_connection: str    # full connection string / URL / path
    location: str          # "Sheet1!A1" | "VBA:Module1:42" | "connections.xml" | ...
    query_text: str | None = None
    parsed_query: ParsedQuery | None = None
    author: str | None = None
    created_at: str | None = None
    modified_at: str | None = None
    metadata: dict = field(default_factory=dict)
    confidence: float = 1.0

    @staticmethod
    def make_id(category: str, raw_connection: str, location: str) -> str:
        h = hashlib.sha256(f"{category}|{raw_connection}|{location}".encode()).hexdigest()[:12]
        return f"{category[:3]}_{h}"

    def to_dict(self) -> dict:
        d = asdict(self)
        return d
