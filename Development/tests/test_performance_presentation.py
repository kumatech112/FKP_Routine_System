"""Synthetic workbook and COM-adapter regressions; no real mail, SQL or Office."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile
import tempfile
import unittest
import xml.etree.ElementTree as ET

from routine.performance_presentation import (
    MAIN, NS, YELLOW, cell_values, highlight_rows, oa15_rows, prepare_mail_file,
    remove_controls, verify_mail_file,
)
from routine.vendor.fkp.domain import ValidationError
from routine.vendor.fkp.workbook import Package, column


def workbook(path, *, wrong_fill=False, amount=1000.25, extras=None):
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    book = ET.Element(f"{{{MAIN}}}workbook")
    sheets = ET.SubElement(book, f"{{{MAIN}}}sheets")
    relations = ET.Element(f"{{{package_rel}}}Relationships")
    for index, name in enumerate(("FC", "SummaryByOACode"), 1):
        ET.SubElement(sheets, f"{{{MAIN}}}sheet", name=name, sheetId=str(index),
                      attrib={f"{{{rel}}}id": f"rId{index}"})
        ET.SubElement(relations, f"{{{package_rel}}}Relationship", Id=f"rId{index}",
                      Type=rel + "/worksheet", Target=f"worksheets/sheet{index}.xml")
    sheet = ET.Element(f"{{{MAIN}}}worksheet")
    data = ET.SubElement(sheet, f"{{{MAIN}}}sheetData")
    row = ET.SubElement(data, f"{{{MAIN}}}row", r="20")
    for index in range(2, 49):
        ref = f"{column(index)}20"
        style = "0" if wrong_fill and ref == "AV20" else "1"
        cell = ET.SubElement(row, f"{{{MAIN}}}c", r=ref, s=style)
        if index in (2, 4):
            ET.SubElement(cell, f"{{{MAIN}}}v").text = str(1 if index == 2 else amount)
        elif index in (3, 5):
            cell.set("t", "inlineStr")
            text = ET.SubElement(cell, f"{{{MAIN}}}is")
            ET.SubElement(text, f"{{{MAIN}}}t").text = "OA 15" if index == 3 else "0000123"
    summary = ET.Element(f"{{{MAIN}}}worksheet")
    ET.SubElement(summary, f"{{{MAIN}}}sheetData")
    styles = (
        f'<styleSheet xmlns="{MAIN}"><fills count="2">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>'
        '</fills><cellXfs count="2"><xf fillId="0"/><xf fillId="1"/></cellXfs></styleSheet>'
    )
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", ET.tostring(book))
        archive.writestr("xl/_rels/workbook.xml.rels", ET.tostring(relations))
        archive.writestr("xl/worksheets/sheet1.xml", ET.tostring(sheet))
        archive.writestr("xl/worksheets/sheet2.xml", ET.tostring(summary))
        archive.writestr("xl/styles.xml", styles)
        archive.writestr("[Content_Types].xml", '<Types xmlns="' + package_rel + '"/>')
        for name, content in (extras or {}).items():
            archive.writestr(name, content)


class Shapes:
    def __init__(self, types):
        self.items = []
        for kind in types:
            shape = SimpleNamespace(Type=kind)
            shape.Delete = lambda s=shape: self.items.remove(s)
            self.items.append(shape)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, index):
        return self.items[index - 1]


class Sheet:
    def __init__(self, types=()):
        self.Shapes = Shapes(types)
        self.areas = {}

    def Range(self, address):
        rule = SimpleNamespace(Interior=SimpleNamespace(), Font=SimpleNamespace())
        rule.SetFirstPriority = lambda: setattr(rule, "first", True)
        def add(**kwargs):
            rule.arguments = kwargs
            return rule
        area = SimpleNamespace(Interior=SimpleNamespace(), Font=SimpleNamespace(),
                               FormatConditions=SimpleNamespace(Add=add), rule=rule)
        self.areas[address] = area
        return area


class Sheets:
    def __init__(self, **items):
        self.items = items

    def __iter__(self):
        return iter(self.items.values())

    def __call__(self, name):
        return self.items[name]


class PresentationTests(unittest.TestCase):
    def test_controls_removed_backwards_including_summary_and_charts_preserved(self):
        product, summary = Sheet((12, 8, 3, 12)), Sheet((8, 13))
        count = remove_controls(SimpleNamespace(Worksheets=Sheets(FC=product, SummaryByOACode=summary)))
        self.assertEqual(count, 4)
        self.assertEqual([s.Type for s in product.Shapes.items], [3])
        self.assertEqual([s.Type for s in summary.Shapes.items], [13])

    def test_all_numbered_oa15_rows_selected_without_matching_oa150_or_headers(self):
        cells = {}
        for rn, number, label in [(5, "No.", "OA 15"), (20, 1, "OA 15"),
                                   (30, 2, "oa15"), (40, 3, "OA\u00a015"),
                                   (50, 4, "OA 150"), (60, 5, "OA 16")]:
            cells[f"B{rn}"] = {"value": number}
            cells[f"C{rn}"] = {"value": label}
        self.assertEqual(oa15_rows(cells), [20, 30, 40])

    def test_yellow_fill_cannot_be_hidden_by_existing_conditional_formatting(self):
        sheet = Sheet()
        highlight_rows(sheet, [20, 30])
        self.assertEqual(set(sheet.areas), {"B20:AV20", "B30:AV30"})
        for area in sheet.areas.values():
            self.assertEqual(area.Interior.Color, YELLOW)
            self.assertEqual(area.Interior.Pattern, 1)
            self.assertTrue(area.Font.Bold)
            self.assertEqual(area.rule.arguments, {"Type": 2, "Formula1": "=TRUE()"})
            self.assertEqual(area.rule.Interior.Color, YELLOW)
            self.assertTrue(area.rule.StopIfTrue)
            self.assertTrue(area.rule.first)

    def test_saved_yellow_checked_through_last_report_column(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "01_Daily_Oct26.xlsx"
            workbook(path)
            verify_mail_file(path, {"FC": [20]})
            workbook(path, wrong_fill=True)
            with self.assertRaisesRegex(ValidationError, "AV20"):
                verify_mail_file(path, {"FC": [20]})

    def test_active_x_parts_and_relationships_block_delivery(self):
        samples = [
            {"xl/activeX/activeX1.bin": b"fixture"},
            {"xl/ctrlProps/ctrlProp1.xml": b"<formControlPr/>"},
            {"xl/_rels/fixture.rels": '<Relationships><Relationship Type="urn/activeXControl"/></Relationships>'},
        ]
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "01_Daily_Oct26.xlsx"
            for extras in samples:
                with self.subTest(part=next(iter(extras))):
                    workbook(path, extras=extras)
                    with self.assertRaises(ValidationError):
                        verify_mail_file(path, {"FC": [20]})

    def test_vml_notes_preserved_but_form_buttons_rejected(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "01_Daily_Oct26.xlsx"
            name = "xl/drawings/vmlDrawing1.vml"
            workbook(path, extras={name: '<xml><ClientData ObjectType="Note"/></xml>'})
            verify_mail_file(path, {"FC": [20]})
            workbook(path, extras={name: '<xml><ClientData ObjectType="Button"/></xml>'})
            with self.assertRaisesRegex(ValidationError, "VML"):
                verify_mail_file(path, {"FC": [20]})

    def test_cell_identity_keeps_leading_zero_identifiers(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "01_Daily_Oct26.xlsx"
            workbook(path)
            values = cell_values(Package(path))
            self.assertEqual(values["FC"]["E20"][0], "0000123")
            self.assertEqual(values["FC"]["D20"][0], 1000.25)
            self.assertNotIn("F20", values["FC"])

    def preparation(self, path, *, change_data=False):
        product, summary = Sheet((12, 3)), Sheet((8,))
        book = SimpleNamespace(Worksheets=Sheets(FC=product, SummaryByOACode=summary))
        book.Save = (lambda: workbook(path, amount=900)) if change_data else lambda: None
        excel = SimpleNamespace(open=lambda _: book, close=lambda b: None)
        manager = SimpleNamespace()
        class Session:
            def __enter__(self):
                return excel
            def __exit__(self, *args):
                pass
        job = SimpleNamespace(check=lambda: None, progress=lambda *args: None)
        return Session(), job, product, summary

    def test_preparation_verifies_saved_report_and_preserves_data(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "01_Daily_Oct26.xlsx"
            workbook(path)
            session, job, product, summary = self.preparation(path)
            with patch("routine.performance_presentation.Excel", return_value=session):
                result = prepare_mail_file(path, ["FC"], job)
            self.assertEqual(result["removed_controls"], 2)
            self.assertEqual(result["oa15_rows"], {"FC": [20]})
            self.assertEqual(product.Shapes.Count, 1)
            self.assertEqual(summary.Shapes.Count, 0)

    def test_unexpected_measure_change_blocks_delivery_after_cleanup(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "01_Daily_Oct26.xlsx"
            workbook(path)
            session, job, *_ = self.preparation(path, change_data=True)
            with patch("routine.performance_presentation.Excel", return_value=session):
                with self.assertRaisesRegex(ValidationError, "เปลี่ยน"):
                    prepare_mail_file(path, ["FC"], job)

    def test_stage05_cannot_be_opened_for_mail_presentation_changes(self):
        with patch("routine.performance_presentation.Excel") as excel:
            with self.assertRaises(ValidationError):
                prepare_mail_file(Path("05_Daily_Oct26.xlsx"), ["FC"], None)
            excel.assert_not_called()


if __name__ == "__main__":
    unittest.main()
