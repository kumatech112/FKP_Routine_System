"""Financial contracts and validation independent of Office and presentation."""
from dataclasses import asdict, dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_DOWN, ROUND_HALF_UP
from pathlib import Path
import calendar
import re

MONTHS = tuple(calendar.month_abbr[i].upper() for i in range(1, 13))
PRODUCTS = ("FC", "LTS", "TES", "BAY", "KAMC", "KHB", "KCC&HP", "CCP", "PWB", "BAYKAMC")
TABLES = {"master": ("dbo_MPFM_Summary_001", "OANAME"), "billcode": ("dbo_MPFM_BillCode_002", "Billcode"), "bucket": ("dbo_MPFM_Bucket_003", "Bucket"), "daily": ("dbo_PFM_DAILY", "OANAME"), "adjust": ("dbo_PFM_ADJUST", "OANAME")}
SHEETS = {"master": "SummaryByOACode", "billcode": "Billcode", "bucket": "SummaryByBucketSystem"}
TEMPLATES = {"master": "Template Master Performance.xlsx", "billcode": "Teamplate Master Performance Billcode.xlsx", "bucket": "Teamplate Master Performance SummaryByBucket.xlsx", "daily": "Daily_Recovery_Performance.xls", "index": "index Bucket.xls"}


class ValidationError(ValueError):
    """A concise message safe to show to the operator."""


class Cancelled(Exception):
    pass


@dataclass(frozen=True, order=True)
class Month:
    year: int
    month: int

    def __post_init__(self):
        if not 2000 <= self.year <= 2099 or not 1 <= self.month <= 12:
            raise ValidationError("เดือนหรือปีไม่ถูกต้อง")

    @property
    def label(self):
        return f"{MONTHS[self.month - 1]}-{self.year % 100:02d}"

    @property
    def first(self):
        return date(self.year, self.month, 1)

    @property
    def next(self):
        return date(self.year + (self.month == 12), self.month % 12 + 1, 1)

    @classmethod
    def parse(cls, label):
        match = re.fullmatch(r"([A-Z]{3})-(\d{2})", str(label).strip().upper())
        if not match or match[1] not in MONTHS:
            raise ValidationError("รูปแบบเดือนต้องเป็น JAN-26")
        return cls(2000 + int(match[2]), MONTHS.index(match[1]) + 1)

    @classmethod
    def from_filename(cls, filename):
        stem = Path(filename).stem.upper()
        found = set()
        for match in re.finditer(r"(?<![A-Z0-9])(" + "|".join(MONTHS) + r")(\d{2})(?![A-Z0-9])", stem):
            found.add(cls(2000 + int(match[2]), MONTHS.index(match[1]) + 1))
        for match in re.finditer(r"(?<![A-Z0-9])(0[1-9]|1[0-2])(\d{2})(?![A-Z0-9])", stem):
            found.add(cls(2000 + int(match[2]), int(match[1])))
        if len(found) != 1:
            raise ValidationError("ชื่อไฟล์ไม่มีเดือนที่ชัดเจน หรือมีหลายเดือน")
        return found.pop()


def number(value, places=None):
    if value is None or isinstance(value, bool):
        raise ValidationError("ข้อมูลตัวเลขว่างหรือไม่ถูกต้อง")
    try:
        result = Decimal(str(value).replace(",", "").strip())
    except InvalidOperation as error:
        raise ValidationError("พบค่าที่ไม่ใช่ตัวเลข") from error
    if not result.is_finite():
        raise ValidationError("พบตัวเลขที่ไม่ถูกต้อง")
    if places == 0 and result != result.to_integral_value():
        raise ValidationError("จำนวนบัญชีต้องเป็นจำนวนเต็ม")
    return result if places is None else result.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def rate(amount, outstanding):
    return (amount / outstanding).quantize(Decimal("0.0001"), rounding=ROUND_DOWN) if outstanding else Decimal(0)


def key(value):
    if isinstance(value, (int, float)) and value == int(value):
        return str(int(value))
    return str(value or "").strip().upper()


def bucket_key(value):
    return re.sub(r"[^A-Z0-9&]", "", key(value).removeprefix("BUCKET"))


def report_date(title, cutoff, month):
    matches = re.findall(r"([A-Za-z]+)\s*,\s*(20\d{2})", str(title))
    valid = [(calendar.month_name[month.month].upper(), str(month.year))]
    if len(matches) != 1 or [(a.upper(), b) for a, b in matches] != valid:
        raise ValidationError("เดือนในหัวรายงานไม่ตรงกับชื่อไฟล์")
    try:
        return date(month.year, month.month, int(number(cutoff, 0)))
    except (ValueError, OverflowError) as error:
        raise ValidationError("วันสิ้นสุดข้อมูลไม่ถูกต้อง") from error


def select_months(available, year, chosen=(), quarter=None):
    wanted = [Month(year, i) for i in range(quarter * 3 - 2, quarter * 3 + 1)] if quarter in (1, 2, 3, 4) else sorted({Month.parse(s) for s in chosen})
    if not wanted or any(m.year != year for m in wanted):
        raise ValidationError("เลือกเดือนในปีเดียวกันอย่างน้อยหนึ่งเดือน")
    missing = {m.label for m in wanted} - set(available)
    if missing:
        raise ValidationError("ไม่มีข้อมูลเดือน " + ", ".join(sorted(missing)))
    return wanted


@dataclass(frozen=True)
class Detail:
    report: str
    month: str
    asof: str
    category: str
    account: int
    outstanding: Decimal
    amount: Decimal
    entity: str = "ALL"
    product: str = "ALL"
    ecm: str = "ALL"
    bucket: str = ""
    efficiency: Decimal | None = None

    def __post_init__(self):
        if self.report not in TABLES or not self.category:
            raise ValidationError("ประเภทข้อมูลหรือรหัสข้อมูลไม่ถูกต้อง")
        month = Month.parse(self.month)
        stamp = date.fromisoformat(self.asof)
        if (stamp.year, stamp.month) != (month.year, month.month):
            raise ValidationError("วันที่ข้อมูลไม่อยู่ในเดือนรายงาน")
        number(self.account, 0)
        number(self.outstanding)
        number(self.amount)

    @property
    def identity(self):
        return (self.report, key(self.month), key(self.entity), key(self.product), key(self.ecm), key(self.bucket), key(self.category))

    @property
    def scope(self):
        return (self.report, key(self.month), key(self.entity), key(self.product), key(self.ecm))

    def json(self):
        return {k: str(v) if isinstance(v, Decimal) else v for k, v in asdict(self).items()}

    @classmethod
    def load(cls, value):
        fields = dict(value)
        for name in ("outstanding", "amount", "efficiency"):
            if fields.get(name) is not None:
                fields[name] = number(fields[name])
        return cls(**fields)


def validate_details(rows):
    if not rows:
        raise ValidationError("ไม่พบข้อมูลที่นำเข้าได้")
    seen, dates = set(), {}
    for row in rows:
        if row.report in ("master", "billcode", "bucket") and key(row.category).startswith(("TOTAL", "GROUP")):
            raise ValidationError("ไม่นำเข้าแถว Group หรือ Total ลงฐานข้อมูล")
        if row.identity in seen:
            raise ValidationError("พบรหัสข้อมูลซ้ำในขอบเขตเดียวกัน")
        seen.add(row.identity)
        if dates.setdefault(row.scope, row.asof) != row.asof:
            raise ValidationError("วันสิ้นสุดข้อมูลในขอบเขตเดียวกันไม่ตรงกัน")
    return rows
