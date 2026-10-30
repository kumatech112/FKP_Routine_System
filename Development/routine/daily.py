"""Daily-only adapter. Mail consumes 01; SQL consumes the verified saved 05."""
from dataclasses import dataclass
from pathlib import Path
import re

from .config import DAILY_PRODUCTS
from .vendor.fkp.daily import run_daily, is_sequence
from .vendor.fkp.database import Database, ComSql
from .vendor.fkp.domain import Detail, Month, ValidationError, key, number, validate_details
from .vendor.fkp.persistence import digest, dump
from .vendor.fkp.workbook import Package


@dataclass(frozen=True)
class DailyArtifacts:
    mail_file: Path
    database_file: Path
    records: tuple
    mail_hash: str
    database_hash: str

    def verify(self):
        if digest(self.mail_file) != self.mail_hash or digest(self.database_file) != self.database_hash:
            raise ValidationError("ไฟล์ขั้น 01 หรือ 05 เปลี่ยนหลังตรวจสอบ")


def records_from_05(path, products=DAILY_PRODUCTS):
    if set(products) - set(DAILY_PRODUCTS) or not products:
        raise ValidationError("รับเฉพาะผลิตภัณฑ์ Daily")
    path = Path(path)
    if not path.name.startswith("05_"):
        raise ValidationError("ลงฐาน Daily ได้จากไฟล์ขั้น 05 เท่านั้น")
    month = Month.from_filename(path)
    package = Package(path)
    if set(package.sheets) != set(products) | {"SummaryByOACode"}:
        raise ValidationError("ชีตในไฟล์ขั้น 05 ไม่ตรงกับ Daily ที่เลือก")
    result = []
    for product in products:
        cells = package.cells(product)
        value = lambda ref: cells.get(ref, {}).get("value")
        row_numbers = sorted({int(re.search(r"\d+$", ref)[0]) for ref in cells})
        for rn in row_numbers:
            if not is_sequence(value(f"B{rn}")):
                continue
            category = key(value(f"C{rn}"))
            if not category or category == "SAM" or category.startswith("TOTAL"):
                raise ValidationError("พบแถวรายละเอียดไม่ถูกต้องในขั้น 05")
            if value(f"H{rn}") != product or not value(f"I{rn}"):
                raise ValidationError("Product/Bucket ในขั้น 05 ไม่ครบ")
            if key(value(f"L{rn}")) != month.label:
                raise ValidationError("เดือนในขั้น 05 ไม่ตรงกับชื่อไฟล์")
            if not re.fullmatch(r"20\d{2}-\d{2}-\d{2}", str(value(f"K{rn}"))):
                raise ValidationError("ไม่พบวันที่ทำงานในขั้น 05")
            result.append(Detail(
                "daily", month.label, str(value(f"J{rn}")), category,
                int(number(value(f"D{rn}"), 0)),
                number(value(f"E{rn}"), 4), number(value(f"F{rn}"), 4),
                product=product, bucket=str(value(f"I{rn}")),
                efficiency=number(value(f"G{rn}"), 4),
            ))
    return tuple(validate_details(result))


def process_daily(source, settings, job):
    expected, outputs = run_daily(source, "daily", list(DAILY_PRODUCTS), settings, job)
    stage01 = job.folder / ("01_" + Path(source).stem + ".xlsx")
    stage05 = job.folder / ("05_" + Path(source).stem + ".xlsx")
    if outputs != [str(stage05)] or not stage01.is_file() or not stage05.is_file():
        raise ValidationError("ผลลัพธ์ Daily ขั้น 01/05 ไม่ครบ")
    before = digest(stage05)
    records = records_from_05(stage05)
    left = {r.identity: r for r in validate_details(expected)}
    right = {r.identity: r for r in records}
    if left != right or digest(stage05) != before:
        raise ValidationError("ข้อมูลที่อ่านจากขั้น 05 ไม่ตรงกับผลประมวลผล Daily")
    result = DailyArtifacts(stage01, stage05, records, digest(stage01), before)
    dump(job.evidence / "daily_routing.json", {
        "mode": "daily", "mail_stage": "01", "database_stage": "05",
        "mail_file": str(stage01), "database_file": str(stage05),
        "mail_sha256": result.mail_hash, "database_sha256": result.database_hash,
        "records": len(records),
    })
    return result


class DailyDatabase(Database):
    """Resolve only dbo_PFM_DAILY; unrelated tables need not exist."""
    def __enter__(self):
        if self.sql is not None:
            self.tables = {"daily": "[dbo].[PFM_DAILY]"}
            return self
        import win32com.client
        dao = win32com.client.Dispatch("DAO.DBEngine.120")
        access = dao.OpenDatabase(str(self.path), False, True)
        try:
            table = access.TableDefs("dbo_PFM_DAILY")
            source, connection = str(table.SourceTableName), str(table.Connect)
            if not re.fullmatch(r"dbo\.[A-Za-z0-9_]+", source) or not connection.startswith("ODBC;"):
                raise ValidationError("ลิงก์ dbo_PFM_DAILY ไม่ถูกต้อง")
            self.tables = {"daily": "[dbo].[" + source.split(".")[1] + "]"}
        finally:
            access.Close()
        self.sql = ComSql(connection)
        return self

    def import_rows(self, rows, replace, job):
        if replace or any(r.report != "daily" for r in rows):
            raise ValidationError("รับเฉพาะ Daily และไม่แทนที่ข้อมูลเดิมอัตโนมัติ")
        return super().import_rows(rows, False, job)
