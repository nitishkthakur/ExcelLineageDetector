"""Convert Excel formulas to SQL-like notation."""
from __future__ import annotations

import re


# Mapping of Excel functions to SQL equivalents
_FUNCTION_MAP = {
    # Aggregation
    "SUM": "SUM",
    "AVERAGE": "AVG",
    "COUNT": "COUNT",
    "COUNTA": "COUNT",
    "COUNTIF": "COUNT_IF",
    "COUNTIFS": "COUNT_IF",
    "SUMIF": "SUM_IF",
    "SUMIFS": "SUM_IF",
    "SUMPRODUCT": "SUM_PRODUCT",
    "MIN": "MIN",
    "MAX": "MAX",
    "MEDIAN": "MEDIAN",
    "STDEV": "STDDEV",
    "STDEVP": "STDDEV_POP",
    "VAR": "VARIANCE",
    "VARP": "VARIANCE_POP",
    # Math
    "ABS": "ABS",
    "ROUND": "ROUND",
    "ROUNDUP": "CEIL",
    "ROUNDDOWN": "FLOOR",
    "INT": "FLOOR",
    "MOD": "MOD",
    "POWER": "POWER",
    "SQRT": "SQRT",
    "LN": "LN",
    "LOG": "LOG",
    "LOG10": "LOG10",
    "EXP": "EXP",
    # String
    "LEFT": "LEFT",
    "RIGHT": "RIGHT",
    "MID": "SUBSTR",
    "LEN": "LENGTH",
    "TRIM": "TRIM",
    "UPPER": "UPPER",
    "LOWER": "LOWER",
    "CONCATENATE": "CONCAT",
    "SUBSTITUTE": "REPLACE",
    "FIND": "INSTR",
    "SEARCH": "INSTR",
    "TEXT": "FORMAT",
    # Logic
    "IF": "CASE WHEN",
    "AND": "AND",
    "OR": "OR",
    "NOT": "NOT",
    "IFERROR": "COALESCE",
    "IFNA": "COALESCE",
    "IFS": "CASE",
    # Lookup
    "VLOOKUP": "JOIN_LOOKUP",
    "HLOOKUP": "JOIN_LOOKUP",
    "INDEX": "INDEX",
    "MATCH": "MATCH",
    "XLOOKUP": "JOIN_LOOKUP",
    "OFFSET": "OFFSET",
    # Date
    "TODAY": "CURRENT_DATE",
    "NOW": "CURRENT_TIMESTAMP",
    "YEAR": "YEAR",
    "MONTH": "MONTH",
    "DAY": "DAY",
    "DATE": "DATE",
    "DATEDIF": "DATEDIFF",
    "EDATE": "DATE_ADD",
    "EOMONTH": "LAST_DAY",
}

# Pattern to match Excel cell references (sheet name cannot contain parens/commas)
_CELL_REF_RE = re.compile(
    r"(?:'?\[?([^\]'!(),]+)\]?'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)"
)
_FUNC_RE = re.compile(r"([A-Z][A-Z0-9_.]+)\s*\(")


def excel_to_sql(formula: str, var_names: dict[str, str] | None = None) -> str:
    """Convert an Excel formula to SQL-like notation.

    Args:
        formula: Excel formula string (without leading =)
        var_names: Optional mapping of cell references to variable names
                   e.g. {"Inputs!B2:B13": "revenue", "Assumptions!B2": "discount_rate"}

    Returns SQL-like representation of the formula.
    """
    if not formula:
        return ""

    # Remove leading = if present
    if formula.startswith("="):
        formula = formula[1:]

    result = formula
    var_names = var_names or {}

    # Replace cell references with variable names where available
    def replace_ref(match: re.Match) -> str:
        sheet = match.group(1) or ""
        ref = match.group(2).replace("$", "")
        full_ref = f"{sheet}!{ref}" if sheet else ref
        # Check for exact match
        if full_ref in var_names:
            return var_names[full_ref]
        # Check ref-only match
        if ref in var_names:
            return var_names[ref]
        # Keep as-is but clean up
        return full_ref if sheet else ref

    result = _CELL_REF_RE.sub(replace_ref, result)

    # Replace Excel functions with SQL equivalents
    def replace_func(match: re.Match) -> str:
        func_name = match.group(1).upper()
        sql_func = _FUNCTION_MAP.get(func_name, func_name)
        return f"{sql_func}("

    result = _FUNC_RE.sub(replace_func, result)

    # Handle IF -> CASE WHEN conversion
    result = _convert_if_to_case(result)

    # Clean up operators
    result = result.replace("<>", "!=")
    result = result.replace("&", " || ")

    return result


def _convert_if_to_case(formula: str) -> str:
    """Convert CASE WHEN(...) to CASE WHEN ... THEN ... ELSE ... END.

    Uses balanced-parenthesis parsing to handle nested function calls
    like CASE WHEN(SUMIF(A:A,">0"), C, D).
    """
    if "CASE WHEN(" not in formula:
        return formula

    result = []
    i = 0
    marker = "CASE WHEN("
    while i < len(formula):
        pos = formula.find(marker, i)
        if pos == -1:
            result.append(formula[i:])
            break

        result.append(formula[i:pos])

        # Find the matching close paren, respecting nesting
        start = pos + len(marker)
        depth = 1
        j = start
        while j < len(formula) and depth > 0:
            if formula[j] == "(":
                depth += 1
            elif formula[j] == ")":
                depth -= 1
            j += 1

        if depth != 0:
            # Unbalanced — leave as-is
            result.append(formula[pos:j])
            i = j
            continue

        inner = formula[start:j - 1]  # content inside CASE WHEN(...)

        # Split on top-level commas (depth=0)
        parts = _split_top_level(inner, ",")
        if len(parts) >= 3:
            cond = parts[0].strip()
            then = parts[1].strip()
            else_ = ",".join(parts[2:]).strip()
            result.append(f"CASE WHEN {cond} THEN {then} ELSE {else_} END")
        elif len(parts) == 2:
            cond = parts[0].strip()
            then = parts[1].strip()
            result.append(f"CASE WHEN {cond} THEN {then} END")
        else:
            result.append(formula[pos:j])

        i = j

    return "".join(result)


def _split_top_level(s: str, sep: str) -> list[str]:
    """Split string on sep, but only at top level (not inside parens)."""
    parts = []
    depth = 0
    start = 0
    for i, ch in enumerate(s):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == sep and depth == 0:
            parts.append(s[start:i])
            start = i + 1
    parts.append(s[start:])
    return parts


def batch_convert(
    formulas: list[dict],
    var_names: dict[str, str] | None = None,
) -> list[dict]:
    """Convert a batch of formula dicts, adding sql_formula field.

    Each dict should have 'formula' key. Returns same dicts with 'sql_formula' added.
    """
    for f in formulas:
        f["sql_formula"] = excel_to_sql(f.get("formula", ""), var_names)
    return formulas
