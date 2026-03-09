"""SQL parser for extracting table/column lineage from SQL statements."""

from __future__ import annotations
import re
from lineage.models import ParsedQuery


def parse(sql: str) -> ParsedQuery | None:
    """Parse SQL and extract tables, columns, joins, and filters.

    Tries sqlglot first, falls back to regex if that fails.
    """
    if not sql or not sql.strip():
        return None

    try:
        import sqlglot
        ast = sqlglot.parse_one(sql, error_level=sqlglot.ErrorLevel.IGNORE)
        if ast is None:
            raise ValueError("sqlglot returned None")

        tables = list({t.name for t in ast.find_all(sqlglot.exp.Table) if t.name})
        columns = list({c.name for c in ast.find_all(sqlglot.exp.Column) if c.name})

        joins = []
        for join in ast.find_all(sqlglot.exp.Join):
            join_type = join.args.get("kind", "")
            join_table = join.this.name if join.this else ""
            join_on = str(join.args.get("on", ""))
            joins.append({"type": str(join_type), "table": join_table, "on": join_on})

        filters = []
        where = ast.find(sqlglot.exp.Where)
        if where:
            filters = [str(where.this)]

        return ParsedQuery(
            tables=tables,
            columns=columns,
            joins=joins,
            filters=filters,
            raw_sql=sql,
        )
    except Exception:
        # Fallback: regex-based extraction
        return _parse_regex(sql)


def _parse_regex(sql: str) -> ParsedQuery:
    """Regex-based SQL parser fallback."""
    tables = re.findall(r'(?i)\b(?:FROM|JOIN)\s+([a-zA-Z_][a-zA-Z0-9_.]*)', sql)
    tables = list(set(t.strip() for t in tables if t.strip()))

    columns = []
    # Try to extract SELECT columns
    select_match = re.search(r'(?i)SELECT\s+(.*?)\s+FROM', sql, re.DOTALL)
    if select_match:
        col_str = select_match.group(1)
        if col_str.strip() != '*':
            raw_cols = col_str.split(',')
            for col in raw_cols:
                col = col.strip()
                # Remove aliases
                col = re.sub(r'(?i)\s+AS\s+\w+$', '', col).strip()
                # Remove table prefix
                if '.' in col:
                    col = col.split('.')[-1]
                if col and not col.startswith('('):
                    columns.append(col)

    return ParsedQuery(
        tables=tables,
        columns=list(set(columns)),
        joins=[],
        filters=[],
        raw_sql=sql,
    )
