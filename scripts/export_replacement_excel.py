#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
导出“数据替换来源表”Excel。

特点：
- 仅使用 Python 标准库，无需安装 openpyxl/xlsxwriter/pandas。
- 输入 JSON，输出 .xlsx。
- 固定表头：序号、原文数据、替换数据、数据来源、来源URL。

用法：
python scripts/export_replacement_excel.py --input records.json --output 数据替换来源表.xlsx
"""

from __future__ import annotations

import argparse
import json
import os
import re
import zipfile
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List
from xml.sax.saxutils import escape

HEADERS = ["序号", "原文数据", "替换数据", "数据来源", "来源URL"]
SHEET_NAME = "数据替换来源表"


def col_letter(index: int) -> str:
    """1-based column index to Excel column letter."""
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    # XML 1.0 illegal control chars except tab/newline/carriage return.
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)
    return text


def cell_xml(row_idx: int, col_idx: int, value: Any, style_id: int = 0) -> str:
    ref = f"{col_letter(col_idx)}{row_idx}"
    text = escape(clean_text(value))
    style = f' s="{style_id}"' if style_id else ""
    return f'<c r="{ref}" t="inlineStr"{style}><is><t>{text}</t></is></c>'


def load_records(path: str) -> List[Dict[str, Any]]:
    with open(path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    if isinstance(payload, dict):
        payload = payload.get("records", [])
    if not isinstance(payload, list):
        raise ValueError("Input JSON must be a list, or an object with a 'records' list.")

    records: List[Dict[str, Any]] = []
    for i, item in enumerate(payload, start=1):
        if not isinstance(item, dict):
            raise ValueError(f"Record {i} is not a JSON object.")
        row = {h: item.get(h, "") for h in HEADERS}
        if not row["序号"]:
            row["序号"] = i
        records.append(row)
    return records


def build_sheet_xml(records: List[Dict[str, Any]]) -> str:
    rows: List[str] = []

    # Header row.
    header_cells = "".join(cell_xml(1, idx, header, style_id=1) for idx, header in enumerate(HEADERS, start=1))
    rows.append(f'<row r="1" ht="24" customHeight="1">{header_cells}</row>')

    # Data rows.
    for r, record in enumerate(records, start=2):
        cells: List[str] = []
        for c, header in enumerate(HEADERS, start=1):
            style_id = 2 if header == "来源URL" and clean_text(record.get(header)) not in ("", "-") else 0
            cells.append(cell_xml(r, c, record.get(header, ""), style_id=style_id))
        rows.append(f'<row r="{r}" ht="42" customHeight="1">{"".join(cells)}</row>')

    dimension = f"A1:E{max(1, len(records) + 1)}"
    sheet_views = (
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        '</sheetView></sheetViews>'
    )
    cols = (
        '<cols>'
        '<col min="1" max="1" width="8" customWidth="1"/>'
        '<col min="2" max="2" width="42" customWidth="1"/>'
        '<col min="3" max="3" width="55" customWidth="1"/>'
        '<col min="4" max="4" width="30" customWidth="1"/>'
        '<col min="5" max="5" width="60" customWidth="1"/>'
        '</cols>'
    )
    sheet_data = f'<sheetData>{"".join(rows)}</sheetData>'

    hyperlinks: List[str] = []
    rels: List[str] = []
    rel_id = 1
    for row_idx, record in enumerate(records, start=2):
        url = clean_text(record.get("来源URL", ""))
        if url and url != "-":
            hyperlinks.append(f'<hyperlink ref="E{row_idx}" r:id="rId{rel_id}"/>')
            rels.append((rel_id, url))
            rel_id += 1

    hyperlinks_xml = f'<hyperlinks>{"".join(hyperlinks)}</hyperlinks>' if hyperlinks else ""
    page_margins = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="{dimension}"/>'
        f'{sheet_views}{cols}{sheet_data}{hyperlinks_xml}{page_margins}'
        '</worksheet>'
    )


def build_sheet_rels_xml(records: List[Dict[str, Any]]) -> str:
    rels: List[str] = []
    rel_id = 1
    for record in records:
        url = clean_text(record.get("来源URL", ""))
        if url and url != "-":
            rels.append(
                f'<Relationship Id="rId{rel_id}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                f'Target="{escape(url)}" TargetMode="External"/>'
            )
            rel_id += 1
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(rels)}'
        '</Relationships>'
    )


def styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Arial"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
    <font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Arial"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E386B"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''


def workbook_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl"/>
  <workbookPr defaultThemeVersion="164011"/>
  <sheets>
    <sheet name="{SHEET_NAME}" sheetId="1" r:id="rId1"/>
  </sheets>
  <calcPr calcId="0"/>
</workbook>'''


def workbook_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def root_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def content_types_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''


def app_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>paragraph-data-update skill</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>{SHEET_NAME}</vt:lpstr></vt:vector></TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>'''


def core_xml() -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>数据替换来源表</dc:title>
  <dc:creator>paragraph-data-update skill</dc:creator>
  <cp:lastModifiedBy>paragraph-data-update skill</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def write_xlsx(records: List[Dict[str, Any]], output_path: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml())
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("docProps/app.xml", app_xml())
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("xl/workbook.xml", workbook_xml())
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml())
        zf.writestr("xl/styles.xml", styles_xml())
        zf.writestr("xl/worksheets/sheet1.xml", build_sheet_xml(records))
        # Hyperlink relationship part is safe even if empty.
        zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", build_sheet_rels_xml(records))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export data replacement records to XLSX.")
    parser.add_argument("--input", "-i", required=True, help="Input JSON file path.")
    parser.add_argument("--output", "-o", required=True, help="Output .xlsx file path.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    records = load_records(args.input)
    write_xlsx(records, args.output)
    print(f"Saved: {args.output}")


if __name__ == "__main__":
    main()
