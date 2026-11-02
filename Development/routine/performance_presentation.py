"""Presentation-only cleanup of the outgoing Performance Daily stage 01."""
from pathlib import Path
import re
import xml.etree.ElementTree as ET

from .vendor.fkp.domain import ValidationError, number
from .vendor.fkp.office import Excel
from .vendor.fkp.workbook import Package, MAIN, NS, column

YELLOW = 65535  # Excel RGB(255, 255, 0)
CONTROL_SHAPES = frozenset((8, 12))  # Form control / ActiveX OLE control
FORM_OBJECTS = frozenset(("button", "checkbox", "radio", "edit", "drop", "list",
                         "group", "spin", "scroll", "label", "dialog"))


def oa15_rows(cells):
    rows = []
    for reference, cell in cells.items():
        if not re.fullmatch(r"C[1-9]\d*", reference):
            continue
        if re.sub(r"\s", "", str(cell["value"] or "")).upper() != "OA15":
            continue
        rn = int(reference[1:])
        try:
            if number(cells.get(f"B{rn}", {}).get("value"), 0) > 0:
                rows.append(rn)
        except ValidationError:
            continue
    return sorted(rows)


def cell_values(package):
    """Ignore style-only blank cells; keep identifiers, measures and formulas exact."""
    return {
        name: {
            ref: (cell["value"], cell["formula"])
            for ref, cell in package.cells(name).items()
            if cell["value"] not in (None, "") or cell["formula"]
        }
        for name in package.sheets
    }


def remove_controls(book):
    removed = 0
    for sheet in book.Worksheets:
        shapes = sheet.Shapes
        for index in range(int(shapes.Count), 0, -1):
            shape = shapes.Item(index)
            if int(shape.Type) in CONTROL_SHAPES:
                shape.Delete()
                removed += 1
        if any(int(shapes.Item(index).Type) in CONTROL_SHAPES
               for index in range(1, int(shapes.Count) + 1)):
            raise ValidationError("ลบ ActiveX หรือปุ่มควบคุมในขั้น 01 ไม่ครบ")
    return removed


def highlight_rows(sheet, rows):
    for rn in rows:
        area = sheet.Range(f"B{rn}:AV{rn}")
        area.Interior.Pattern = 1  # xlSolid
        area.Interior.Color = YELLOW
        area.Font.Bold = True
        # Keep existing conditional rules. This row-only rule ensures their fill
        # cannot conceal OA15 yellow; other rows keep their relative rule order.
        rule = area.FormatConditions.Add(Type=2, Formula1="=TRUE()")
        rule.Interior.Pattern = 1
        rule.Interior.Color = YELLOW
        rule.Font.Bold = True
        rule.StopIfTrue = True
        rule.SetFirstPriority()


def verify_mail_file(path, expected_rows):
    package = Package(path)
    for name, (_, content) in package.parts.items():
        lower = name.lower()
        if lower.startswith(("xl/activex/", "xl/ctrlprops/")) or lower == "xl/vbaproject.bin":
            raise ValidationError("ขั้น 01 ยังมี ActiveX หรือส่วนควบคุม จึงยังไม่ส่งเมล")
        if lower.endswith((".rels", ".vml")) or name == "[Content_Types].xml":
            root = ET.fromstring(content)
            for node in root.iter():
                kind = node.attrib.get("Type", node.attrib.get("ContentType", "")).lower()
                if "activex" in kind or kind.endswith(("/ctrlprop", "/control")):
                    raise ValidationError("ขั้น 01 ยังอ้างถึง ActiveX หรือส่วนควบคุม")
                if node.tag.rsplit("}", 1)[-1] == "ClientData":
                    if node.attrib.get("ObjectType", "").lower() in FORM_OBJECTS:
                        raise ValidationError("ขั้น 01 ยังมีปุ่มควบคุมใน VML")
    styles = ET.fromstring(package.parts["xl/styles.xml"][1])
    fills = list(styles.find("s:fills", NS))
    formats = list(styles.find("s:cellXfs", NS))
    for name, rows in expected_rows.items():
        sheet = ET.fromstring(package.parts[package.sheets[name]][1])
        if any(node.tag.rsplit("}", 1)[-1] == "controls" for node in sheet.iter()):
            raise ValidationError("ขั้น 01 ยังมีส่วนควบคุมบนชีต")
        indexed = {int(row.attrib["r"]): row
                   for row in sheet.findall("s:sheetData/s:row", NS)}
        for rn in rows:
            row = indexed.get(rn)
            cells = {} if row is None else {c.attrib["r"]: c for c in row}
            for col in range(2, 49):
                ref = f"{column(col)}{rn}"
                cell = cells.get(ref)
                style = int(cell.attrib.get("s", "0")) if cell is not None else 0
                fill = fills[int(formats[style].attrib.get("fillId", "0"))]
                pattern = fill.find("s:patternFill", NS)
                foreground = None if pattern is None else pattern.find("s:fgColor", NS)
                if (pattern is None or pattern.attrib.get("patternType") != "solid"
                        or foreground is None
                        or foreground.attrib.get("rgb", "").upper()[-6:] != "FFFF00"):
                    raise ValidationError("สีเหลือง OA15 ในขั้น 01 ไม่ครบ: " + name + "!" + ref)
    return package


def prepare_mail_file(path, products, job):
    path = Path(path)
    if not path.name.startswith("01_") or path.suffix.lower() != ".xlsx":
        raise ValidationError("ปรับรูปแบบได้เฉพาะไฟล์ Performance ขั้น 01")
    before = Package(path)
    if set(before.sheets) != set(products) | {"SummaryByOACode"}:
        raise ValidationError("ชีตขั้น 01 ไม่ตรงกับ Daily ที่เลือก")
    expected = {name: oa15_rows(before.cells(name)) for name in products}
    original_values = cell_values(before)
    job.check()
    job.progress(88, "ลบ ActiveX และใส่สีเหลือง OA15 ในไฟล์ขั้น 01")
    with Excel() as excel:
        book = excel.open(path)
        removed = remove_controls(book)
        for name, rows in expected.items():
            job.check()
            highlight_rows(book.Worksheets(name), rows)
        job.check()
        book.Save()
        excel.close(book)
    after = verify_mail_file(path, expected)
    if list(before.sheets) != list(after.sheets) or cell_values(after) != original_values:
        raise ValidationError("ค่าหรือสูตรในขั้น 01 เปลี่ยนระหว่างปรับรูปแบบ จึงยังไม่ส่งเมล")
    return {"removed_controls": removed, "oa15_rows": expected}
