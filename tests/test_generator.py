"""Test workbook generator for Excel Lineage Detector tests.

Creates an Excel file with planted connections of various types to test
detection coverage.
"""

from __future__ import annotations
import io
import json
import zipfile
from pathlib import Path


def generate_test_workbook(path: Path) -> list[dict]:
    """Generate a test workbook with planted connections.

    Creates an xlsx file with as many planted connection types as possible
    using openpyxl for base creation plus direct ZIP/XML manipulation for
    connection types that openpyxl doesn't support natively.

    Args:
        path: Output path for the generated xlsx file.

    Returns:
        List of planted connections: [{"id": ..., "type": ..., "description": ...}]
    """
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font

    # Step 1: Create base workbook with openpyxl
    wb = Workbook()
    ws_main = wb.active
    ws_main.title = "Sheet1"

    # Plant some basic content
    ws_main["A1"] = "External Reference Formula (planted below via XML)"
    ws_main["B1"] = "This sheet has planted external data connections"
    ws_main["C1"] = "WEBSERVICE Formula"
    ws_main["A2"] = "Normal data"
    ws_main["B2"] = 42
    ws_main["C2"] = "=TODAY()"  # Normal formula

    # Add a visible hyperlink via openpyxl
    ws_main["A5"] = "External URL link"
    ws_main["A5"].hyperlink = "https://example.com/data-source"
    ws_main["A5"].font = Font(color="0000FF", underline="single")

    # Add another hyperlink to a file
    ws_main["A6"] = "External file link"
    ws_main["A6"].hyperlink = "file:///C:/data/source_data.xlsx"
    ws_main["A6"].font = Font(color="0000FF", underline="single")

    # Add hidden sheet with formula placeholder
    ws_hidden = wb.create_sheet("HiddenData")
    ws_hidden.sheet_state = "hidden"
    ws_hidden["A1"] = "Hidden sheet external ref placeholder"

    # Add a comment with URL
    from openpyxl.comments import Comment
    comment = Comment(
        "Data sourced from https://api.data.gov/economic/v1/stats and "
        "also from \\\\fileserver\\reports\\quarterly.xlsx",
        "DataEngineer"
    )
    ws_main["D1"] = "Has comment with URL"
    ws_main["D1"].comment = comment

    # Save to buffer first
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Step 2: Read all zip contents
    zip_contents: dict[str, bytes] = {}
    with zipfile.ZipFile(buf, "r") as zf:
        for name in zf.namelist():
            zip_contents[name] = zf.read(name)

    # Step 3: Inject XML modifications

    # 3a. Modify Sheet1 XML to add external formula references directly
    sheet1_xml = zip_contents.get("xl/worksheets/sheet1.xml", b"")
    if sheet1_xml:
        sheet1_xml = _inject_external_formulas(sheet1_xml)
        zip_contents["xl/worksheets/sheet1.xml"] = sheet1_xml

    # 3b. Modify hidden sheet XML to add external formula
    hidden_sheet_key = None
    for k in zip_contents:
        if "worksheets/sheet" in k and k.endswith(".xml"):
            if k != "xl/worksheets/sheet1.xml":
                hidden_sheet_key = k
                break
    if hidden_sheet_key:
        hidden_xml = zip_contents[hidden_sheet_key]
        hidden_xml = _inject_hidden_sheet_formula(hidden_xml)
        zip_contents[hidden_sheet_key] = hidden_xml

    # 3c. Inject connections.xml with ODBC, OLEDB, and web query connections
    connections_xml = _make_connections_xml()
    zip_contents["xl/connections.xml"] = connections_xml.encode("utf-8")

    # 3d. Update workbook.xml to reference connections and add definedNames
    wb_xml = zip_contents.get("xl/workbook.xml", b"")
    if wb_xml:
        wb_xml = _inject_workbook_additions(wb_xml)
        zip_contents["xl/workbook.xml"] = wb_xml

    # 3e. Inject Power Query custom XML
    pq_item_xml = _make_power_query_xml()
    zip_contents["xl/customXml/item0.xml"] = pq_item_xml.encode("utf-8")
    pq_props_xml = _make_custom_xml_props()
    zip_contents["xl/customXml/itemProps0.xml"] = pq_props_xml.encode("utf-8")

    # 3f. Add custom document properties with file path
    custom_props_xml = _make_custom_doc_props()
    zip_contents["docProps/custom.xml"] = custom_props_xml.encode("utf-8")

    # 3g. Update [Content_Types].xml to include new files
    content_types = zip_contents.get("[Content_Types].xml", b"")
    content_types = _update_content_types(content_types)
    zip_contents["[Content_Types].xml"] = content_types

    # Step 4: Write new zip
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in zip_contents.items():
            zf.writestr(name, data)

    # Return list of planted connections
    planted = [
        {
            "id": "ext_formula_xlsx",
            "type": "external_workbook",
            "description": "Cell formula referencing [source_data.xlsx]Sheet1'!A1",
        },
        {
            "id": "ext_formula_unc",
            "type": "unc_path",
            "description": "Cell formula with UNC path \\\\fileserver\\share\\[quarterly.xlsx]",
        },
        {
            "id": "webservice_formula",
            "type": "webservice",
            "description": "WEBSERVICE formula calling https://api.exchangerate.host/latest",
        },
        {
            "id": "odbc_connection",
            "type": "odbc",
            "description": "ODBC connection to SQL Server via connections.xml",
        },
        {
            "id": "oledb_connection",
            "type": "oracle",
            "description": "OLE DB connection to Oracle via connections.xml (detected as oracle sub-type)",
        },
        {
            "id": "web_query_connection",
            "type": "web",
            "description": "Web query connection via connections.xml",
        },
        {
            "id": "power_query_m",
            "type": "sql_server",
            "description": "Power Query M code connecting to SQL Server",
        },
        {
            "id": "hyperlink_url",
            "type": "http",
            "description": "External URL hyperlink to https://example.com/data-source",
        },
        {
            "id": "hyperlink_file",
            "type": "file_url",
            "description": "External file hyperlink to file:///C:/data/source_data.xlsx",
        },
        {
            "id": "comment_url",
            "type": "comment_url",
            "description": "URL in cell comment: https://api.data.gov/economic/v1/stats",
        },
        {
            "id": "comment_unc",
            "type": "unc_path",
            "description": "UNC path in cell comment: \\\\fileserver\\reports\\quarterly.xlsx",
        },
        {
            "id": "custom_prop_path",
            "type": "custom_property_file_path",
            "description": "File path in custom document property",
        },
        {
            "id": "named_range_external",
            "type": "named_range_external",
            "description": "Named range referencing external workbook [budget.xlsx]",
        },
        {
            "id": "hidden_sheet_formula",
            "type": "external_workbook",
            "description": "External formula in hidden sheet",
        },
        {
            "id": "metadata_core",
            "type": "core_properties",
            "description": "Document core metadata (creator, dates)",
        },
    ]

    return planted


def _inject_external_formulas(sheet_xml: bytes) -> bytes:
    """Inject external formula references into sheet XML."""
    xml_str = sheet_xml.decode("utf-8", errors="replace")

    # Find the sheetData section and add formula rows
    # We insert rows with external formulas after existing rows
    formula_rows = """
    <row r="10">
      <c r="A10" t="str">
        <f>'[source_data.xlsx]Sheet1'!A1</f>
        <v></v>
      </c>
    </row>
    <row r="11">
      <c r="A11" t="str">
        <f>'\\\\fileserver\\share\\[quarterly.xlsx]Q1'!B2</f>
        <v></v>
      </c>
    </row>
    <row r="12">
      <c r="A12" t="str">
        <f>WEBSERVICE("https://api.exchangerate.host/latest")</f>
        <v></v>
      </c>
    </row>
"""

    # Insert before </sheetData>
    if "</sheetData>" in xml_str:
        xml_str = xml_str.replace("</sheetData>", formula_rows + "</sheetData>")
    else:
        # Try to find sheetData and append
        if "<sheetData/>" in xml_str:
            xml_str = xml_str.replace(
                "<sheetData/>",
                "<sheetData>" + formula_rows + "</sheetData>"
            )

    return xml_str.encode("utf-8")


def _inject_hidden_sheet_formula(sheet_xml: bytes) -> bytes:
    """Inject external formula into hidden sheet."""
    xml_str = sheet_xml.decode("utf-8", errors="replace")

    formula_row = """
    <row r="1">
      <c r="A1" t="str">
        <f>'[external_source.xlsx]DataSheet'!C5</f>
        <v></v>
      </c>
    </row>
"""

    if "</sheetData>" in xml_str:
        xml_str = xml_str.replace("</sheetData>", formula_row + "</sheetData>")
    elif "<sheetData/>" in xml_str:
        xml_str = xml_str.replace(
            "<sheetData/>",
            "<sheetData>" + formula_row + "</sheetData>"
        )

    return xml_str.encode("utf-8")


def _make_connections_xml() -> str:
    """Create connections.xml with ODBC, OLE DB, and web query connections."""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <connection id="1" name="SalesDB_ODBC" type="1" refreshedVersion="3" background="1" savePassword="0">
    <dbPr connection="Driver={SQL Server};Server=sql-server.corp.local;Database=SalesDB;Trusted_Connection=Yes;"
          command="SELECT * FROM dbo.SalesFact WHERE Year = 2023"
          commandType="2"/>
  </connection>
  <connection id="2" name="OracleFinance_OleDB" type="4" refreshedVersion="3" background="0" savePassword="0">
    <dbPr connection="Provider=OraOLEDB.Oracle;Data Source=oracle-db.corp.local:1521/FINDB;User Id=finance_user;Password=***;"
          command="SELECT account_id, amount FROM gl_balances WHERE period_year = 2023"
          commandType="2"/>
  </connection>
  <connection id="3" name="ExternalWebQuery" type="3" refreshedVersion="2" background="0" savePassword="0">
    <webPr url="https://data.worldbank.org/api/v2/country/US/indicator/NY.GDP.MKTP.CD?format=json"
           xml="0" sourceData="1" parsePre="0" sequential="0" firstRow="0"
           xl97="0" textDates="0" htmlTables="0" post="0" htmlFormat="none"/>
  </connection>
  <connection id="4" name="Query - SalesQuery" type="5" refreshedVersion="3" background="0" savePassword="0">
    <dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=SalesQuery;Extended Properties=&quot;&quot;;"
          command="SELECT * FROM [SalesQuery]"
          commandType="2"/>
  </connection>
</connections>
'''


def _inject_workbook_additions(wb_xml: bytes) -> bytes:
    """Add definedNames with external reference to workbook.xml."""
    xml_str = wb_xml.decode("utf-8", errors="replace")

    defined_names_xml = """
  <definedNames>
    <definedName name="ExternalBudget">'[budget.xlsx]Summary'!$A$1:$Z$100</definedName>
    <definedName name="LocalRange">Sheet1!$A$1:$D$10</definedName>
  </definedNames>
"""

    # Insert before </workbook>
    if "</workbook>" in xml_str:
        # Replace self-closing <definedNames/> if present, otherwise inject before </workbook>
        if "<definedNames/>" in xml_str:
            xml_str = xml_str.replace("<definedNames/>", defined_names_xml.strip())
        elif "<definedNames>" not in xml_str:
            xml_str = xml_str.replace("</workbook>", defined_names_xml + "</workbook>")

    return xml_str.encode("utf-8")


def _make_power_query_xml() -> str:
    """Create customXml/item0.xml with Power Query M code."""
    pq_content = {
        "Queries": [
            {
                "Name": "SalesData",
                "Description": "Loads sales data from SQL Server",
                "Formula": 'let\n    Source = Sql.Database("sql-server.corp.local", "SalesDB"),\n    dbo_SalesFact = Source{[Schema="dbo",Item="SalesFact"]}[Data],\n    #"Filtered Rows" = Table.SelectRows(dbo_SalesFact, each [Year] = 2023)\nin\n    #"Filtered Rows"',
                "IsParameterQuery": False,
                "Type": "Table",
            },
            {
                "Name": "ExchangeRates",
                "Description": "Fetches live exchange rates from web API",
                "Formula": 'let\n    Source = Web.Contents("https://api.exchangerate.host/latest"),\n    ParsedJson = Json.Document(Source)\nin\n    ParsedJson',
                "IsParameterQuery": False,
                "Type": "Table",
            },
            {
                "Name": "LocalFileData",
                "Description": "Reads data from a local CSV file",
                "Formula": 'let\n    Source = Csv.Document(File.Contents("C:\\\\data\\\\input\\\\sales_export.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]),\n    #"Promoted Headers" = Table.PromoteHeaders(Source)\nin\n    #"Promoted Headers"',
                "IsParameterQuery": False,
                "Type": "Table",
            },
        ]
    }
    json_str = json.dumps(pq_content, indent=2)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbookPr xmlns:x="urn:schemas-microsoft-com:office:excel">
<![CDATA[{json_str}]]>
</x:workbookPr>
'''


def _make_custom_xml_props() -> str:
    """Create itemProps0.xml for the custom XML item."""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ds:datastoreItem ds:itemID="{A6140000-1111-2222-3333-44445555AAAB}"
  xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
  <ds:schemaRefs/>
</ds:datastoreItem>
'''


def _make_custom_doc_props() -> str:
    """Create docProps/custom.xml with file path property."""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="DataSourcePath">
    <vt:lpwstr>\\\\fileserver\\shared\\datasets\\master_data.xlsx</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="ReportingServer">
    <vt:lpwstr>https://reporting.corp.local/api/v2/reports</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Author">
    <vt:lpwstr>DataEngineer</vt:lpwstr>
  </property>
</Properties>
'''


def _update_content_types(content_types: bytes) -> bytes:
    """Update [Content_Types].xml to include custom XML and connections."""
    xml_str = content_types.decode("utf-8", errors="replace")

    # Add content types if not present
    additions = []

    if "customXml/item0.xml" not in xml_str:
        additions.append(
            '<Override PartName="/xl/customXml/item0.xml" '
            'ContentType="application/xml"/>'
        )

    if "connections.xml" not in xml_str:
        additions.append(
            '<Override PartName="/xl/connections.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.connections+xml"/>'
        )

    if "docProps/custom.xml" not in xml_str:
        additions.append(
            '<Override PartName="/docProps/custom.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'custom-properties+xml"/>'
        )

    if additions and "</Types>" in xml_str:
        xml_str = xml_str.replace(
            "</Types>",
            "\n".join(additions) + "\n</Types>"
        )

    return xml_str.encode("utf-8")
