#!/usr/bin/env python3
import argparse
import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COL_RE = re.compile(r"([A-Z]+)")


def col_to_index(ref):
    match = COL_RE.match(ref or "")
    if not match:
        return None
    value = 0
    for char in match.group(1):
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def index_to_col(index):
    out = []
    while index:
        index, rem = divmod(index - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def text_of(node):
    if node is None:
        return ""
    return "".join(part.text or "" for part in node.iter() if part.tag.endswith("}t"))


def first_sheet_path(zf):
    workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
    first_sheet = workbook_root.find(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet")
    if first_sheet is None:
        raise ValueError("workbook has no sheets")
    rel_id = first_sheet.attrib.get(f"{{{REL_NS}}}id")
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib["Target"]
            if target.startswith("/"):
                return target.lstrip("/")
            return str(Path("xl") / target)
    raise ValueError(f"relationship not found for {rel_id}")


def shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [text_of(si) for si in root.findall(f"{{{MAIN_NS}}}si")]


def cell_value(cell, strings):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        inline = cell.find(f"{{{MAIN_NS}}}is")
        return text_of(inline)
    value = cell.find(f"{{{MAIN_NS}}}v")
    raw = value.text if value is not None and value.text is not None else ""
    if cell_type == "s":
        return strings[int(raw)] if raw else ""
    if cell_type == "b":
        return "TRUE" if raw == "1" else "FALSE"
    return raw


def iter_rows(path):
    with zipfile.ZipFile(path) as zf:
        sheet_path = first_sheet_path(zf)
        strings = shared_strings(zf)
        with zf.open(sheet_path) as xml_file:
            for _event, row in ET.iterparse(xml_file, events=("end",)):
                if row.tag != f"{{{MAIN_NS}}}row":
                    continue
                values = []
                fallback_col = 1
                for cell in row.findall(f"{{{MAIN_NS}}}c"):
                    col_index = col_to_index(cell.attrib.get("r")) or fallback_col
                    while len(values) < col_index - 1:
                        values.append("")
                    values.append(cell_value(cell, strings))
                    fallback_col = col_index + 1
                yield values
                row.clear()


def normalize_row(row, width):
    row = list(row[:width])
    if len(row) < width:
        row.extend([""] * (width - len(row)))
    return row


def discover(files):
    header = None
    width = 0
    data_rows = 0
    per_file = []
    mismatched_headers = []
    for path in files:
        row_count = 0
        file_header = None
        for row in iter_rows(path):
            if row_count == 0:
                file_header = row
                if header is None:
                    header = row
                    width = len(header)
                elif normalize_row(row, width) != normalize_row(header, width):
                    mismatched_headers.append(path.name)
            else:
                data_rows += 1
            row_count += 1
        per_file.append({"file": path.name, "rows": max(row_count - 1, 0)})
    if header is None:
        raise ValueError("no rows found in source workbooks")
    return header, width, data_rows, per_file, mismatched_headers


def write_cell(out, row_index, col_index, value, style):
    if value is None or value == "":
        return
    ref = f"{index_to_col(col_index)}{row_index}"
    safe = escape(str(value))
    out.write(
        f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">'
        f"{safe}</t></is></c>"
    )


def write_row(out, row_index, values, width, style):
    out.write(f'<row r="{row_index}">')
    for col_index, value in enumerate(normalize_row(values, width), start=1):
        write_cell(out, row_index, col_index, value, style)
    out.write("</row>")


def worksheet_xml(files, header, width, total_rows, sheet_path):
    last_col = index_to_col(width)
    column_widths = [10, 10, 8, 14, 10, 10, 22, 28, 14, 28, 18, 8, 12]
    with sheet_path.open("w", encoding="utf-8") as out:
        out.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        out.write(f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">')
        out.write(f'<dimension ref="A1:{last_col}{total_rows}"/>')
        out.write(
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            "</sheetView></sheetViews>"
        )
        out.write('<sheetFormatPr baseColWidth="8" defaultRowHeight="15"/>')
        out.write("<cols>")
        for idx in range(1, width + 1):
            width_value = column_widths[idx - 1] if idx <= len(column_widths) else 16
            out.write(f'<col min="{idx}" max="{idx}" width="{width_value}" customWidth="1"/>')
        out.write("</cols><sheetData>")
        write_row(out, 1, header, width, 1)
        out_row = 2
        for path in files:
            for input_row, row in enumerate(iter_rows(path)):
                if input_row == 0:
                    continue
                write_row(out, out_row, row, width, 2)
                out_row += 1
        out.write("</sheetData>")
        out.write(f'<autoFilter ref="A1:{last_col}{total_rows}"/>')
        out.write('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        out.write("</worksheet>")


def write_package(output_path, sheet_xml_path):
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""",
        "xl/workbook.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="OCR Results" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        "xl/styles.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{MAIN_NS}">
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0F766E"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="49" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"/>
    <xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>""",
        "docProps/core.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>""",
        "docProps/app.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>OCR Results</vt:lpstr></vt:vector></TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>""",
    }
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
        zf.write(sheet_xml_path, "xl/worksheets/sheet1.xml")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("source_dir")
    parser.add_argument("output_xlsx")
    args = parser.parse_args()

    source_dir = Path(args.source_dir).expanduser()
    output_path = Path(args.output_xlsx).expanduser()
    files = sorted(source_dir.glob("*.xlsx"))
    if not files:
        raise SystemExit(f"No .xlsx files found in {source_dir}")

    header, width, data_rows, per_file, mismatched_headers = discover(files)
    total_rows = data_rows + 1
    if total_rows > 1_048_576:
        raise SystemExit(f"Merged result has {total_rows} rows, over Excel's row limit")

    tmpdir = Path(tempfile.mkdtemp(prefix="merge_xlsx_"))
    try:
        sheet_path = tmpdir / "sheet1.xml"
        worksheet_xml(files, header, width, total_rows, sheet_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_output = output_path.with_suffix(output_path.suffix + ".tmp")
        write_package(tmp_output, sheet_path)
        shutil.move(tmp_output, output_path)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    summary = {
        "source_dir": str(source_dir),
        "output": str(output_path),
        "files": len(files),
        "columns": width,
        "data_rows": data_rows,
        "total_rows_with_header": total_rows,
        "mismatched_headers": mismatched_headers,
        "first_file": per_file[0],
        "last_file": per_file[-1],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
