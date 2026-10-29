"""Independent five-stage Daily/Adjust processor driven by workbook contracts."""
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re

from .domain import Detail, Month, PRODUCTS, ValidationError, bucket_key, key, number, report_date, validate_details
from .office import CacheFilters, Excel, matrix
from .persistence import dump, now
from .workbook import Package


def is_sequence(value):
    if value is None or isinstance(value, bool):
        return False
    try:
        return number(value, 0) > 0
    except ValidationError:
        return False


def numeric_or_placeholder(value):
    if value in (None, "", "-"):
        return Decimal(0)
    return number(value)


def is_oa15(value):
    return re.sub(r"\s", "", key(value)) == "OA15"


@dataclass
class Block:
    start: int
    end: int
    title: str
    header: int
    data: list[int]
    total: int


def blocks(grid, strict=True):
    starts = [i for i, row in enumerate(grid) if key(row[0]).startswith("BUCKET")]
    result = []
    for start, end in zip(starts, starts[1:] + [len(grid)]):
        header = next((i for i in range(start, end) if key(grid[i][1]).rstrip(".") == "NO" and key(grid[i][2]) == "NAME"), None)
        total = next((i for i in range(start, end) if key(grid[i][2]).startswith("TOTAL")), None)
        if header is None or total is None or total <= header:
            if strict:
                raise ValidationError("โครงสร้างตาราง Bucket ไม่ครบ")
            continue
        data = []
        for i in range(header + 1, total):
            if is_sequence(grid[i][1]):
                data.append(i)
            else:
                break
        result.append(Block(start, end, str(grid[start][0]), header, data, total))
    return result


def cleanup_rows(grid):
    remove = set()
    for block in blocks(grid):
        remove.update(i for i in block.data if key(grid[i][2]) in ("", "SAM"))
        remove.update(i for i in block.data if len(grid[i]) > 6 and isinstance(grid[i][6], (int, float)) and grid[i][6] == 0)
        for i in range(block.header + 1, block.total):
            row = grid[i]
            if key(row[2]) == "SAM" and key(row[3]) == "SAM" and "SAM" in key(row[4]):
                remove.add(i)
    return sorted(remove)


def blacklist(title):
    normalized = bucket_key(title)
    return normalized.startswith(("TOTALALL", "TOTALLEGALMEASUREMENT", "TOTALLEGALNONTYPE", "LEGBBMEASUREMENT", "LEGBILL66", "LEGNOTBILL66", "TOTALLEGALTYPE"))


def sort_key(row):
    units, outstanding = numeric_or_placeholder(row[3]), numeric_or_placeholder(row[4])
    if row[7] in (None, "", "-"):
        return 3, -outstanding, Decimal(0)
    efficiency = number(row[7])
    return (1 if units > 100 else 2), -efficiency, -units


def rank_block(grid, block):
    eligible = all(any(key(grid[i][c]).rstrip(".") in words for i in range(block.start, block.header + 1)) for c, words in ((7, {"EFF"}), (8, {"STANDARD"}), (9, {"SCORE"}), (10, {"RANK"})))
    if not eligible or not block.data:
        return {}
    sorted_rows = sorted((list(grid[i]) for i in block.data), key=sort_key)
    changes = {}
    for number_in_order, (index, row) in enumerate(zip(block.data, sorted_rows), 1):
        row[1] = number_in_order
        row[10] = number_in_order if numeric_or_placeholder(row[3]) > 100 and row[7] not in (None, "", "-") else "-"
        changes[index] = row
    return changes


def process_plan(grid, mode):
    remove, edits = set(), {}
    for block in blocks(grid):
        oa15 = next((i for i in block.data if is_oa15(grid[i][2])), None)
        collected = any(numeric_or_placeholder(grid[i][5]) != 0 for i in block.data)
        keep = oa15 is not None and collected and not blacklist(block.title)
        if mode == "adjust" and oa15 is not None:
            keep = keep and numeric_or_placeholder(grid[oa15][5]) > 0
        if not keep:
            remove.update(range(block.start, block.end))
            continue
        if mode == "daily" and grid[oa15][5] in (None, "", "-"):
            edited = list(grid[oa15])
            edited[5] = 0
            edits[oa15] = edited
            grid = list(grid)
            grid[oa15] = tuple(edited)
        edits.update(rank_block(grid, block))
        if bucket_key(block.title) == "WO0":
            candidates = [i for i in range(block.start + 1, block.header) if key(grid[i][2]) == "OA"]
            if candidates:
                remove.add(candidates[0])
    return sorted(remove), edits


def delete_rows(sheet, indexes):
    groups = []
    for index in sorted(set(indexes)):
        rn = index + 1
        if groups and rn == groups[-1][1] + 1:
            groups[-1][1] = rn
        else:
            groups.append([rn, rn])
    for first, last in reversed(groups):
        sheet.Rows(f"{first}:{last}").Delete()


def last_row(sheet):
    last = sheet.Range("A:AV").Find(What="*", After=sheet.Range("A1"), LookIn=-4123, LookAt=2, SearchOrder=1, SearchDirection=2, MatchCase=False, SearchFormat=False)
    return int(last.Row) if last is not None else 1


def grid_for(sheet, last_column="L"):
    return matrix(sheet, f"A1:{last_column}{last_row(sheet)}")


def snapshot(book, path):
    book.SaveAs(str(path), FileFormat=51)
    return path


def details_for(sheet, product, mode, month, stamp, mapping):
    grid = grid_for(sheet)
    result = []
    for block in blocks(grid):
        code = mapping.get(bucket_key(block.title))
        if code is None:
            raise ValidationError("ไม่พบรหัส Bucket " + block.title)
        for index in block.data:
            row = grid[index]
            if key(row[2]) in ("", "SAM"):
                raise ValidationError("ยังมีชื่อว่างหรือ SAM ในรายงาน")
            amount = numeric_or_placeholder(row[5])
            if mode == "adjust" and is_oa15(row[2]) and amount <= 0:
                raise ValidationError("OA 15 ไม่มี Collected ในรายงานตัวจบ")
            result.append(Detail(mode, month.label, stamp, key(row[2]), int(number(row[3], 0)), number(row[4], 4), number(amount, 4), product=product, bucket=code, efficiency=number(numeric_or_placeholder(row[7]), 4)))
    return result


def merged_metrics(row, partner=None):
    values = [numeric_or_placeholder(row[c]) + (numeric_or_placeholder(partner[c]) if partner else 0) for c in (3, 4, 5)]
    efficiency = Decimal(0) if values[0] < 100 or not values[1] else (values[2] / values[1] * 100).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return values, efficiency


def combine_bay_kamc(book, selected, month, stamp, mapping):
    bay, kamc = book.Worksheets("BAY"), book.Worksheets("KAMC")
    target = book.Worksheets("BAYKAMC")
    grid, other = grid_for(target), grid_for(kamc)
    other_blocks = {bucket_key(b.title): b for b in blocks(other)}
    remove = []
    for block in blocks(grid):
        counterpart = other_blocks.get(bucket_key(block.title))
        if counterpart is None:
            remove.extend(range(block.start, block.end))
            continue
        partners = {re.sub(r"\s|(?<=OA)0+(?=\d)", "", key(other[i][2])): other[i] for i in counterpart.data}
        for i in block.data:
            row = grid[i]
            partner = partners.get(re.sub(r"\s|(?<=OA)0+(?=\d)", "", key(row[2])))
            values, efficiency = merged_metrics(row, partner)
            target.Range(f"D{i+1}:F{i+1}").Value2 = (tuple(float(v) for v in values),)
            target.Range(f"H{i+1}").Value2 = float(efficiency)
        paired_rows = [(block.total, counterpart.total)]
        inhouse = next((i for i in range(block.header + 1, block.total) if key(grid[i][2]) == "INHOUSE"), None)
        other_inhouse = next((i for i in range(counterpart.header + 1, counterpart.total) if key(other[i][2]) == "INHOUSE"), None)
        if inhouse is not None:
            paired_rows.append((inhouse, other_inhouse))
        for index, other_index in paired_rows:
            values, efficiency = merged_metrics(grid[index], other[other_index] if other_index is not None else None)
            target.Range(f"D{index+1}:F{index+1}").Value2 = (tuple(float(v) for v in values),)
            target.Range(f"H{index+1}").Value2 = float(efficiency)
    delete_rows(target, remove)
    return target


def run_daily(source_path, mode, products, settings, job):
    if mode not in ("daily", "adjust") or not products or set(products) - set(PRODUCTS):
        raise ValidationError("เลือกงานและผลิตภัณฑ์ให้ถูกต้อง")
    if mode == "daily" and "BAYKAMC" in products:
        raise ValidationError("BAYKAMC ใช้ได้เฉพาะตัวจบ")
    month = Month.from_filename(source_path)
    staged = job.stage(source_path)
    template = job.stage(settings["templates"]["daily"])
    index = job.stage(settings["templates"]["index"])
    sources = [p for p in PRODUCTS if p != "BAYKAMC" and (p in products or "BAYKAMC" in products and p in ("BAY", "KAMC"))]
    paths = []
    with Excel() as excel:
        source, book, index_book = excel.open(staged), excel.open(template), excel.open(index, readonly=True)
        mapping = {bucket_key(a): str(b) for a, b in matrix(index_book.Worksheets(1), "A2:B100") if a and b}
        stamp = report_date(source.Worksheets("SummaryByOACode").Range("A1").Value2, source.Worksheets("SummaryByOACode").Range("C3").Value2, month).isoformat()
        summary = source.Worksheets("SummaryByOACode")
        summary.Range("A1:BF43").Copy(Destination=book.Worksheets("SummaryByOACode").Range("A1"))
        book.Worksheets("SummaryByOACode").Range("A1:BF43").Value2 = summary.Range("A1:BF43").Value2
        filters = CacheFilters(source, "Data_Per")
        native = source.Worksheets("PerformanceByProductBySystem")
        for index_number, product in enumerate(sources):
            job.check()
            target = book.Worksheets(product)
            entity, group, product_filter = [row[0] for row in matrix(target, "C5:C7")]
            filters.apply({"Entity": entity, "Product_Group": ["2IN1", "3IN1"] if group == "2IN1&3IN1" else group, "Product": product_filter}, job)
            native.Range("C5").Value2, native.Range("C6").Value2, native.Range("C7").Value2 = entity, group, product_filter
            excel.calculate()
            end = last_row(native)
            if end < 22:
                raise ValidationError("ไม่พบตารางรายละเอียดในต้นทาง")
            area = f"A1:AV{end}"
            target.Cells.ClearContents()
            native.Range(area).Copy(Destination=target.Range("A1"))
            target.Range(area).Value2 = native.Range(area).Value2
            delete_rows(target, cleanup_rows(grid_for(target, "AV")))
            job.progress(5 + 25 * (index_number + 1) / len(sources), "เตรียมข้อมูล " + product)
        for sheet in list(book.Worksheets):
            if sheet.Name not in sources + ["SummaryByOACode"]:
                sheet.Delete()
        for step in range(1, 6):
            job.check()
            if step == 2:
                for product in sources:
                    sheet = book.Worksheets(product)
                    for columns in ("J:AO", "H:H", "D:F"):
                        sheet.Columns(columns).Delete()
            elif step == 3:
                for product in sources:
                    sheet = book.Worksheets(product)
                    remove, edits = process_plan(grid_for(sheet), mode)
                    for index_row, row in edits.items():
                        sheet.Range(f"A{index_row+1}:L{index_row+1}").Value2 = (tuple(row),)
                    delete_rows(sheet, remove)
                    retained = blocks(grid_for(sheet))
                    if retained and retained[0].start > 2:
                        delete_rows(sheet, list(range(2, retained[0].start)))
            elif step == 4:
                if "BAYKAMC" in products:
                    original_names = {s.Name for s in book.Worksheets}
                    current = Path(book.FullName)
                    excel.close(book)
                    package = Package(current)
                    package.clone_sheet("BAY", "BAYKAMC")
                    cloned = job.folder / "adjust_baykamc.xlsx"
                    package.save(cloned)
                    book = excel.open(cloned)
                    if {s.Name for s in book.Worksheets} != original_names | {"BAYKAMC"}:
                        raise ValidationError("ตรวจสอบการเพิ่มชีต BAYKAMC ไม่ผ่าน")
                    combine_bay_kamc(book, products, month, stamp, mapping)
                records = []
                for product in products:
                    records.extend(details_for(book.Worksheets(product), product, mode, month, stamp, mapping))
                validate_details(records)
                for product in products:
                    sheet = book.Worksheets(product)
                    sheet.Columns("I:L").Delete()
                    sheet.Columns("G:G").Delete()
                    if product == "BAYKAMC":
                        for block in blocks(grid_for(sheet)):
                            for i in block.data + [block.total]:
                                sheet.Range(f"G{i+1}").Formula = f"=IF(D{i+1}<100,0,IFERROR(F{i+1}/E{i+1},0))"
                                sheet.Range(f"G{i+1}").NumberFormat = "0.00%"
            elif step == 5:
                for product in products:
                    sheet = book.Worksheets(product)
                    # Routine adapter: keep identifiers/dates/month labels literal in saved 05.
                    sheet.Range("H:L").NumberFormat = "@"
                    grid = grid_for(sheet)
                    for block in blocks(grid):
                        for i in block.data:
                            if is_oa15(grid[i][2]):
                                sheet.Range(f"B{i+1}:L{i+1}").Font.Bold = True
                        header = next((i for i in range(block.start, block.header) if key(grid[i][2]) == "CURRENT MONTH"), None)
                        if header is not None:
                            sheet.Range(f"G{header+1}:L{header+1}").Value2 = (("EFF.", "PRODUCT", "BUCKET", "DATE_STAMP", "ADD_DATE", "MONTH_n"),)
                        for i in block.data:
                            sheet.Range(f"H{i+1}:L{i+1}").Value2 = ((product, mapping[bucket_key(block.title)], stamp, now().date().isoformat(), month.label),)
                    sheet.Range("A:L").Columns.AutoFit()
                for sheet in list(book.Worksheets):
                    if sheet.Name not in list(products) + ["SummaryByOACode"]:
                        sheet.Delete()
            path = job.folder / (f"{step:02d}_" + Path(source_path).stem + ".xlsx")
            excel.calculate()
            snapshot(book, path)
            paths.append(path)
            job.progress(30 + step * 11, "บันทึกรายงานขั้นที่ " + str(step))
    job.verify_sources()
    dump(job.evidence / "daily_check.json", {"mode": mode, "products": products, "records": len(records), "steps": [str(p) for p in paths]})
    return records, [str(paths[-1])]
