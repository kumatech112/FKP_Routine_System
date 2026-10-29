"""Fresh report readers with independent control totals and exact filter scopes."""
from decimal import Decimal

from .domain import Detail, Month, SHEETS, ValidationError, key, number, report_date, validate_details
from .office import CacheFilters, Excel, matrix
from .workbook import Package


def close_measure(actual, expected):
    return abs(number(actual) - number(expected)) <= max(Decimal("0.02"), abs(number(expected)) * Decimal("0.0000000001"))


def check_control(rows, control, included=None):
    sample = [r for r in rows if included is None or r.category in included]
    totals = (sum(r.account for r in sample), sum(r.outstanding for r in sample), sum(r.amount for r in sample))
    if any(not close_measure(a, b) for a, b in zip(totals, control)):
        raise ValidationError("ยอดรวมรายละเอียดไม่ตรงกับต้นทาง")
    return [str(n) for n in totals]


def saved_master(path, month):
    cells = Package(path).cells(SHEETS["master"])
    value = lambda ref: cells.get(ref, {}).get("value")
    stamp = report_date(value("A1"), value("C3"), month).isoformat()
    rows, control = [], None
    for rn in range(14, 150):
        category = key(value(f"B{rn}"))
        if not category:
            continue
        measures = [value(f"{c}{rn}") for c in "CDE"]
        if category == "TOTAL":
            control = measures
            break
        rows.append(Detail("master", month.label, stamp, category, int(number(measures[0], 0)), number(measures[1], 2), number(measures[2], 2)))
    if control is None:
        raise ValidationError("ไม่พบยอดรวมต้นทาง SummaryByOACode")
    validate_details(rows)
    check_control(rows, control)
    return rows


def summary_rows(book, report, month, profile, ecm):
    sheet = book.Worksheets(SHEETS[report])
    stamp = report_date(sheet.Range("A1").Value2, sheet.Range("C3").Value2, month).isoformat()
    rows, control = [], None
    for raw in matrix(sheet, "B14:E80"):
        category = key(raw[0])
        if not category:
            continue
        if category == "TOTAL":
            control = raw[1:4]
            break
        rows.append(Detail(report, month.label, stamp, category, int(number(raw[1], 0)), number(raw[2], 2), number(raw[3], 2), key(profile["entity"]), key(profile["product"]), ecm))
    if control is None:
        raise ValidationError("ไม่พบยอดรวมต้นทาง " + SHEETS[report])
    included = {str(i) for i in range(60, 70)} | {"OTHER"} if report == "billcode" else None
    totals = check_control(rows, control, included)
    return rows, totals


def extract_source(path, reports, settings, job, start=0, end=65):
    month = Month.from_filename(path)
    staged = job.stage(path)
    result = []
    controls = []
    profiles = [p for p in settings["profiles"] if p["enabled"]]
    work = [(r, p, e) for r in reports if r != "master" for p in profiles for e in (("ALL", "YES", "NO") if r == "billcode" else ("ALL",))]
    count = len(work) + ("master" in reports)
    done = 0
    if "master" in reports:
        result.extend(saved_master(staged, month))
        done += 1
        job.progress(start + (end - start) * done / count, "ตรวจสอบ SummaryByOACode")
    if work:
        with Excel() as excel:
            book = excel.open(staged)
            filters = CacheFilters(book, "Data_Bucket")
            for report, profile, ecm in work:
                job.check()
                filters.apply({"OA_Code": settings["oa_code"], "Entity": profile["entity"], "Product": profile["product"], "Product_Group": profile["product_group"], "ECM": ecm}, job)
                sheet = book.Worksheets(SHEETS[report])
                for address, val in {"C5": settings["oa_code"], "C6": profile["entity"], "G3": profile["product"], "G4": profile["product_group"], "K7": ecm}.items():
                    sheet.Range(address).Value2 = "(All)" if key(val) == "ALL" else val
                excel.calculate()
                rows, totals = summary_rows(book, report, month, profile, ecm)
                result.extend(rows)
                controls.append({"report": report, "sheet": profile["sheet"], "ecm": ecm, "totals": totals})
                done += 1
                job.progress(start + (end - start) * done / count, "อ่าน " + profile["sheet"] + " " + ecm)
    job.verify_sources()
    validate_details(result)
    return result, controls
