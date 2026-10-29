from pathlib import Path
from typing import Any
import re
STABLE_TEXT_FORMAT = "@"
STABLE_DATE_FORMAT = "dd/mm/yyyy"
STABLE_DATETIME_FORMAT = "dd/mm/yyyy hh:mm:ss"
STABLE_TIME_FORMAT = "hh:mm:ss"
STABLE_DURATION_FORMAT = "[h]:mm:ss"
STABLE_INTEGER_FORMAT = "#,##0"
STABLE_DECIMAL_FORMAT = "#,##0.00"
STABLE_DEFAULT_FONT = "Tahoma"

def normalize_stable_header(value: Any) -> str:
    return re.sub(r"[^0-9a-zก-๙]+", "", str(value or "").strip().lower())


def is_duration_header(header: Any) -> bool:
    key = normalize_stable_header(header)
    raw = str(header or "").strip().lower()
    return any(
        marker in key or marker in raw
        for marker in (
            "talktime",
            "talktimes",
            "totaltalktimes",
            "avgtt",
            "firstcall",
            "duration",
            "ระยะเวลา",
        )
    )


def is_date_header(header: Any) -> bool:
    key = normalize_stable_header(header)
    raw = str(header or "").strip().lower()
    if is_duration_header(header):
        return False
    return any(marker in key or marker in raw for marker in ("date", "วันที่", "paydate", "postingdate"))


def is_time_header(header: Any) -> bool:
    key = normalize_stable_header(header)
    raw = str(header or "").strip().lower()
    if is_duration_header(header) or is_date_header(header):
        return False
    return any(marker in key or marker in raw for marker in ("time", "เวลา"))


def is_decimal_header(header: Any) -> bool:
    key = normalize_stable_header(header)
    raw = str(header or "").strip().lower()
    return any(
        marker in key or marker in raw
        for marker in (
            "amount",
            "balance",
            "payamount",
            "paymentamount",
            "sumofpaymentamount",
            "บาท",
            "ยอด",
        )
    )


def is_integer_header(header: Any) -> bool:
    key = normalize_stable_header(header)
    raw = str(header or "").strip().lower()
    if is_duration_header(header):
        return False
    return any(
        marker in key or marker in raw
        for marker in (
            "answer",
            "noanswer",
            "noanswered",
            "count",
            "total",
            "sumof",
            "รวมcall",
            "จำนวน",
        )
    )


def is_text_header(header: Any) -> bool:
    key = normalize_stable_header(header)
    raw = str(header or "").strip().lower()
    if not key:
        return False
    if key in {"no", "rowno", "total"}:
        return False
    if is_duration_header(header) or is_date_header(header) or is_time_header(header):
        return False
    if is_decimal_header(header) or is_integer_header(header):
        return False
    return any(
        marker in key or marker in raw
        for marker in (
            "agreement",
            "agreementn",
            "agreemen",
            "wkagno",
            "account",
            "phone",
            "tel",
            "mobile",
            "caller",
            "callee",
            "callerid",
            "channel",
            "wpclno",
            "wkclno",
            "clno",
            "tcc",
            "tccc",
            "collector",
            "staff",
            "agent",
            "idcard",
            "idstaff",
            "staffid",
            "staffcode",
            "product",
            "code",
            "resultcode",
            "actioncode",
            "typejob",
            "jobno",
            "bucket",
            "name",
            "รหัส",
            "เบอร์",
            "โทร",
            "สัญญา",
            "ชื่อ",
            "เลขที่",
            "เลือกตั้ง",
        )
    )


def stable_excel_value(header: Any, value: Any) -> Any:
    if value is None:
        return "" if is_text_header(header) else None
    if is_text_header(header):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        if isinstance(value, int):
            return str(value)
        return str(value).strip()
    return value


def infer_stable_number_format(header: Any, explicit_format: str | None = None) -> str | None:
    if explicit_format:
        return explicit_format
    if is_text_header(header):
        return STABLE_TEXT_FORMAT
    if is_duration_header(header):
        return STABLE_DURATION_FORMAT
    if is_date_header(header):
        return STABLE_DATETIME_FORMAT
    if is_time_header(header):
        return STABLE_TIME_FORMAT
    if is_decimal_header(header):
        return STABLE_DECIMAL_FORMAT
    if is_integer_header(header):
        return STABLE_INTEGER_FORMAT
    return None


def apply_stable_openpyxl_sheet_formats(
    sheet: Any,
    *,
    header_row: int = 1,
    data_start_row: int | None = None,
    max_row: int | None = None,
    font_name: str | None = STABLE_DEFAULT_FONT,
) -> None:
    from copy import copy

    data_start_row = data_start_row or (header_row + 1)
    max_row = max_row or sheet.max_row

    for column_index in range(1, sheet.max_column + 1):
        header = sheet.cell(header_row, column_index).value
        number_format = infer_stable_number_format(header)
        text_column = is_text_header(header)
        for row_index in range(header_row, max_row + 1):
            cell = sheet.cell(row_index, column_index)
            if font_name and cell.value is not None:
                font = copy(cell.font)
                font.name = font_name
                cell.font = font
            if row_index < data_start_row:
                continue
            if text_column:
                cell.value = stable_excel_value(header, cell.value)
                if isinstance(cell.value, str):
                    cell.data_type = "s"
            if number_format:
                cell.number_format = number_format


def open_dao_database(path):
    import win32com.client
    engine = win32com.client.Dispatch("DAO.DBEngine.120")
    return engine, engine.OpenDatabase(str(path), False, False)


def save_openpyxl_workbook(book, path, overwrite=True):
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Payment output must be XLSX")
    if path.exists() and not overwrite:
        raise FileExistsError(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)


def required_string(values, name, section):
    value = values.get(name)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(section + "." + name)
    return value


def optional_string(value):
    return str(value).strip() if value is not None and str(value).strip() else None


def resolve_path(config, value):
    path = Path(value)
    return path if path.is_absolute() else Path(config).parent / path
