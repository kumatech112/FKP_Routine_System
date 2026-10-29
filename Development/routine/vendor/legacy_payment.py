from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook

from . import payment_support as base


IMPORT_HEADERS = [
    "WKOACD",
    "WKPDTH",
    "WKPCUT",
    "AgreementN",
    "Payment_date",
    "Posting_Date",
    "Payment_Amount",
    "System_Time",
    "WKOBAL",
    "WKBCKT",
    "WKLTYP",
    "WKCLNO",
    "Cycle_Month",
    "Reassign_Date",
    "Round",
]

SOURCE_HEADERS = [
    "WKOACD",
    "WKPDTH",
    "WKPCUT",
    "WKAGNO",
    "Payment_date",
    "Posting_Date",
    "Payment_Amount",
    "System_Time",
    "WKOBAL",
    "WKBCKT",
    "WKLTYP",
    "WKCLNO",
    "Cycle_Month",
    "Reassign_Date",
]

OUTPUT_EXTRA_HEADERS = [
    "TcccNo",
    "collectorname",
    "TYPE_JOB",
    "เลือกตั้งNo",
    "FORECAST",
]

ACTION_DEDUPLICATION_KEY_COLUMNS = ("AgreementN", "COLLCETION", "type")


@dataclass(frozen=True)
class PaymentOnlineConfig:
    access_database_path: Path
    source_workbook_path: Path
    source_sheet_name: str | None
    output_path: Path
    import_sheet_name: str
    query_sheet_name: str
    overwrite: bool
    run_action_queries: bool = True


@dataclass(frozen=True)
class SourceWorkbookData:
    headers: list[str]
    rows: list[list[Any]]
    import_rows: list[list[Any]]
    round_label: str


@dataclass(frozen=True)
class ActionInsertPlan:
    target_table: str
    target_columns: list[str]
    select_sql: str


@dataclass(frozen=True)
class ActionQueryResult:
    query_name: str
    target_table: str
    selected_rows: int
    deleted_existing_rows: int
    inserted_rows: int
    skipped_duplicate_source_rows: int


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Import Payment Online Excel data into Access, update Table1.Round using Query2, "
            "run the Payment Online action queries, and export the final workbook."
        )
    )
    parser.add_argument("--config", required=True, help="Path to the JSON configuration file.")
    parser.add_argument("--source-workbook", help="Optional source Excel workbook path override.")
    parser.add_argument("--output", help="Optional output workbook path override.")
    parser.add_argument("--source-sheet", help="Optional source worksheet name override.")
    return parser.parse_args()


def load_config(config_path: Path, args: argparse.Namespace) -> PaymentOnlineConfig:
    raw = json.loads(config_path.read_text(encoding="utf-8"))
    access_raw = raw.get("access", {})
    source_raw = raw.get("source_workbook", {})
    workbook_raw = raw.get("output_workbook", {})

    source_workbook_path = base.resolve_path(
        config_path,
        base.required_string(source_raw, "path", "source_workbook"),
    )
    if args.source_workbook:
        source_workbook_path = Path(args.source_workbook).expanduser().resolve()

    output_path = base.resolve_path(
        config_path,
        base.required_string(workbook_raw, "output_path", "output_workbook"),
    )
    if args.output:
        output_path = Path(args.output).expanduser().resolve()

    return PaymentOnlineConfig(
        access_database_path=base.resolve_path(
            config_path,
            base.required_string(access_raw, "database_path", "access"),
        ),
        source_workbook_path=source_workbook_path,
        source_sheet_name=args.source_sheet or base.optional_string(source_raw.get("sheet_name")),
        output_path=output_path,
        import_sheet_name=base.optional_string(workbook_raw.get("import_sheet_name")) or "Table1",
        query_sheet_name=base.optional_string(workbook_raw.get("query_sheet_name")) or "Query1",
        overwrite=bool(workbook_raw.get("overwrite", True)),
    )


def extract_source_workbook_data(workbook_path: Path, sheet_name: str | None) -> SourceWorkbookData:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_name] if sheet_name and sheet_name in workbook.sheetnames else workbook[workbook.sheetnames[0]]

        rows = list(worksheet.iter_rows(values_only=True))
        if not rows:
            raise ValueError(f"Source workbook is empty: {workbook_path}")

        raw_headers = [normalize_header_name(value) for value in rows[0]]
        expected_headers = SOURCE_HEADERS[:]
        if raw_headers[: len(expected_headers)] != expected_headers:
            raise ValueError(
                "Source workbook headers do not match the expected Payment Online layout.\n"
                f"Expected: {expected_headers}\n"
                f"Actual:   {raw_headers[:len(expected_headers)]}"
            )

        round_label = extract_round_label_from_filename(workbook_path.name)

        source_rows: list[list[Any]] = []
        import_rows: list[list[Any]] = []
        for raw_row in rows[1:]:
            values = list(raw_row[: len(expected_headers)])
            if row_is_empty(values):
                continue
            source_rows.append(values)

            import_values = values[:]
            import_values[3] = values[3]
            import_values.append(round_label)
            import_rows.append(import_values)

        return SourceWorkbookData(
            headers=expected_headers,
            rows=source_rows,
            import_rows=import_rows,
            round_label=round_label,
        )
    finally:
        workbook.close()


def normalize_header_name(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def row_is_empty(values: Iterable[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True


def extract_round_label_from_filename(file_name: str) -> str:
    match = re.search(r"(?:^|[_-])R(\d+)(?:[_-]|\.|$)", file_name, re.IGNORECASE)
    if not match:
        raise ValueError(f"Unable to determine Payment Online round from file name: {file_name}")
    return f"Payment Online Round {int(match.group(1))}"


def create_working_database_copy(source_database_path: Path, temp_dir: Path) -> Path:
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", source_database_path.stem).strip("._")
    if not safe_name:
        safe_name = "payment_online"
    working_database_path = temp_dir / f"{safe_name}_run{source_database_path.suffix or '.mdb'}"
    working_database_path.write_bytes(source_database_path.read_bytes())
    return working_database_path


def create_import_workbook(workbook_path: Path, rows: list[list[Any]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Table1"
    worksheet.append(IMPORT_HEADERS)
    for row in rows:
        worksheet.append(row)
    base.apply_stable_openpyxl_sheet_formats(worksheet, header_row=1)
    workbook.save(workbook_path)


def import_rows_into_access(database_path: Path, rows: list[list[Any]]) -> None:
    _engine, database = base.open_dao_database(database_path)
    recordset = None
    try:
        database.Execute("DELETE FROM [Table1]")
        recordset = database.OpenRecordset("Table1")
        for row in rows:
            recordset.AddNew()
            for field_name, value in zip(IMPORT_HEADERS, row):
                recordset.Fields(field_name).Value = dao_import_value(value)
            recordset.Update()
    finally:
        if recordset is not None:
            try:
                recordset.Close()
            except Exception:
                pass
        database.Close()


def dao_import_value(value: Any) -> Any:
    if value == "":
        return None
    return value


def fetch_query_rows(database_path: Path, query_name: str, *, order_by: str | None = None) -> tuple[list[str], list[list[Any]]]:
    _engine, database = base.open_dao_database(database_path)
    try:
        source_name = f"[{query_name}]"
        if order_by:
            source_name = f"(SELECT * FROM [{query_name}] ORDER BY {order_by})"
        recordset = database.OpenRecordset(f"SELECT * FROM {source_name}")
        try:
            headers = [recordset.Fields(index).Name for index in range(recordset.Fields.Count)]
            rows: list[list[Any]] = []
            while not recordset.EOF:
                rows.append([recordset.Fields(index).Value for index in range(recordset.Fields.Count)])
                recordset.MoveNext()
            return headers, rows
        finally:
            recordset.Close()
    finally:
        database.Close()


def build_round_updates(query_headers: list[str], query_rows: list[list[Any]], round_label: str) -> list[tuple[int, str]]:
    index_by_name = {header: index for index, header in enumerate(query_headers)}
    no_index = index_by_name.get("No")

    if no_index is None:
        raise ValueError("Query2 is missing expected field: No")

    updates: list[tuple[int, str]] = []
    for row in query_rows:
        updates.append((int(row[no_index]), round_label))
    return updates


def build_excel_text_values(query_headers: list[str], query_rows: list[list[Any]]) -> list[list[str]]:
    index_by_name = {header: index for index, header in enumerate(query_headers)}
    tccc_index = resolve_query_index(index_by_name, "TcccNo", fallback=2)
    type_job_index = resolve_query_index(index_by_name, "TYPE_JOB1", fallback=11)
    election_index = resolve_query_index(index_by_name, "จ่าย/ไม่จ่ายวิชาชีพ", fallback=28)
    forecast_index = resolve_query_index(index_by_name, "FORECAST", fallback=29)

    collector_indexes = [
        index_by_name.get("collectorname1"),
        index_by_name.get("collectorname2"),
        index_by_name.get("ชื่อพนักงาน"),
        index_by_name.get("dbo_รหัสพนักงาน GE.collectorname"),
        index_by_name.get("รหัสพนักงาน GE.collectorname"),
    ]

    text_values: list[list[str]] = []
    for row in query_rows:
        collector_name = first_non_empty(
            row[index]
            for index in collector_indexes
            if index is not None and index < len(row)
        )
        text_values.append([
            normalize_excel_text(row[tccc_index]),
            normalize_excel_text(collector_name),
            normalize_excel_text(row[type_job_index]),
            normalize_excel_text(row[election_index]),
            normalize_excel_text(row[forecast_index]),
        ])
    return text_values


def resolve_query_index(index_by_name: dict[str, int], field_name: str, *, fallback: int) -> int:
    index = index_by_name.get(field_name)
    if index is not None:
        return index
    return fallback


def first_non_empty(values: Iterable[Any]) -> Any:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def normalize_excel_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def apply_round_updates(database_path: Path, updates: list[tuple[int, str]]) -> None:
    _engine, database = base.open_dao_database(database_path)
    try:
        for row_no, round_text in updates:
            escaped = round_text.replace("'", "''")
            database.Execute(f"UPDATE [Table1] SET [Round]='{escaped}' WHERE [No]={row_no}")
    finally:
        database.Close()


def run_action_queries(database_path: Path) -> list[ActionQueryResult]:
    _engine, database = base.open_dao_database(database_path)
    executed: list[ActionQueryResult] = []
    try:
        query_names = [database.QueryDefs(index).Name for index in range(database.QueryDefs.Count)]
        action_query_names = [
            name
            for name in query_names
            if "Payment Online" in name and "Action" in name
        ]
        for query_name in action_query_names:
            executed.append(
                execute_action_query_idempotently(
                    database,
                    query_name,
                    str(database.QueryDefs(query_name).SQL),
                )
            )
    finally:
        database.Close()
    return executed


def execute_action_query_idempotently(database: Any, query_name: str, sql: str) -> ActionQueryResult:
    plan = parse_action_insert_sql(sql)
    insert_rows = fetch_action_insert_rows(database, plan)
    unique_rows, skipped_duplicate_source_rows = unique_action_rows(insert_rows)

    deleted_existing_rows = 0
    inserted_rows = 0
    for row in unique_rows:
        deleted_existing_rows += delete_existing_action_rows(database, plan, row)
        inserted_rows += insert_action_row(database, plan, row)

    return ActionQueryResult(
        query_name=query_name,
        target_table=plan.target_table,
        selected_rows=len(insert_rows),
        deleted_existing_rows=deleted_existing_rows,
        inserted_rows=inserted_rows,
        skipped_duplicate_source_rows=skipped_duplicate_source_rows,
    )


def parse_action_insert_sql(sql: str) -> ActionInsertPlan:
    normalized_sql = sql.strip().rstrip(";").strip()
    match = re.match(
        r"INSERT\s+INTO\s+(?P<target>\[[^\]]+\]|[^\s(]+)\s*"
        r"\((?P<columns>.*?)\)\s*(?P<select>SELECT\s+.*)$",
        normalized_sql,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not match:
        raise ValueError(f"Unsupported Payment Online action query SQL: {sql[:300]}")

    target_table = strip_identifier(match.group("target"))
    target_columns = [strip_identifier(part) for part in split_sql_list(match.group("columns"))]
    if not target_columns:
        raise ValueError(f"Payment Online action query has no target columns: {sql[:300]}")

    return ActionInsertPlan(
        target_table=target_table,
        target_columns=target_columns,
        select_sql=match.group("select").strip(),
    )


def split_sql_list(text: str) -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    bracket_depth = 0
    for char in text:
        if char == "[":
            bracket_depth += 1
        elif char == "]" and bracket_depth:
            bracket_depth -= 1
        elif char == "," and bracket_depth == 0:
            part = "".join(current).strip()
            if part:
                parts.append(part)
            current = []
            continue
        current.append(char)

    part = "".join(current).strip()
    if part:
        parts.append(part)
    return parts


def strip_identifier(identifier: str) -> str:
    value = identifier.strip()
    if value.startswith("[") and value.endswith("]"):
        return value[1:-1].replace("]]", "]")
    return value


def fetch_action_insert_rows(database: Any, plan: ActionInsertPlan) -> list[dict[str, Any]]:
    recordset = database.OpenRecordset(plan.select_sql)
    try:
        field_count = recordset.Fields.Count
        if field_count < len(plan.target_columns):
            raise ValueError(
                f"Action query select returns {field_count} fields, "
                f"but target has {len(plan.target_columns)} columns."
            )

        rows: list[dict[str, Any]] = []
        while not recordset.EOF:
            row: dict[str, Any] = {}
            for index, column_name in enumerate(plan.target_columns):
                row[column_name] = recordset.Fields(index).Value
            rows.append(row)
            recordset.MoveNext()
        return rows
    finally:
        recordset.Close()


def unique_action_rows(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    unique_rows: list[dict[str, Any]] = []
    seen_keys: set[tuple[Any, ...]] = set()
    skipped = 0

    for row in rows:
        key = tuple(normalize_action_key_value(row.get(column_name)) for column_name in ACTION_DEDUPLICATION_KEY_COLUMNS)
        if key in seen_keys:
            skipped += 1
            continue
        seen_keys.add(key)
        unique_rows.append(row)

    return unique_rows, skipped


def normalize_action_key_value(value: Any) -> Any:
    if value is None:
        return None
    return str(value).strip()


def delete_existing_action_rows(database: Any, plan: ActionInsertPlan, row: dict[str, Any]) -> int:
    where_sql = build_action_key_where_clause(row)
    database.Execute(f"DELETE FROM {quote_identifier(plan.target_table)} WHERE {where_sql}")
    return int(getattr(database, "RecordsAffected", 0) or 0)


def insert_action_row(database: Any, plan: ActionInsertPlan, row: dict[str, Any]) -> int:
    column_sql = ", ".join(quote_identifier(column_name) for column_name in plan.target_columns)
    value_sql = ", ".join(format_access_sql_value(row.get(column_name)) for column_name in plan.target_columns)
    database.Execute(f"INSERT INTO {quote_identifier(plan.target_table)} ({column_sql}) VALUES ({value_sql})")
    return int(getattr(database, "RecordsAffected", 0) or 0)


def build_action_key_where_clause(row: dict[str, Any]) -> str:
    return " AND ".join(
        build_access_sql_comparison(column_name, row.get(column_name))
        for column_name in ACTION_DEDUPLICATION_KEY_COLUMNS
    )


def build_access_sql_comparison(column_name: str, value: Any) -> str:
    column_sql = quote_identifier(column_name)
    if value is None:
        return f"{column_sql} IS NULL"
    return f"{column_sql}={format_access_sql_value(value)}"


def quote_identifier(identifier: str) -> str:
    return f"[{identifier.replace(']', ']]')}]"


def format_access_sql_value(value: Any) -> str:
    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return "-1" if value else "0"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, dt.datetime):
        return f"#{value:%m/%d/%Y %H:%M:%S}#"
    if isinstance(value, dt.date):
        return f"#{value:%m/%d/%Y}#"
    escaped = str(value).replace("'", "''")
    return f"'{escaped}'"


def build_output_rows(source_rows: list[list[Any]], excel_text_values: list[list[str]]) -> list[list[Any]]:
    output_rows: list[list[Any]] = []
    for index, source_row in enumerate(source_rows):
        extra_values = excel_text_values[index] if index < len(excel_text_values) else [""] * len(OUTPUT_EXTRA_HEADERS)
        output_rows.append(source_row + extra_values)
    return output_rows


def write_output_workbook(
    output_path: Path,
    *,
    import_sheet_name: str,
    query_sheet_name: str,
    source_headers: list[str],
    source_rows: list[list[Any]],
    query_headers: list[str],
    query_rows: list[list[Any]],
    overwrite: bool,
) -> None:
    if output_path.exists() and not overwrite:
        raise ValueError(f"Output workbook already exists: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    import_sheet = workbook.active
    import_sheet.title = import_sheet_name
    import_sheet.append(source_headers + OUTPUT_EXTRA_HEADERS)
    for row in source_rows:
        import_sheet.append(row)
    import_sheet.freeze_panes = "A2"
    import_sheet.auto_filter.ref = import_sheet.dimensions

    query_sheet = workbook.create_sheet(query_sheet_name)
    query_sheet.append(query_headers)
    for row in query_rows:
        query_sheet.append(row)
    query_sheet.freeze_panes = "A2"
    query_sheet.auto_filter.ref = query_sheet.dimensions
    query_sheet.sheet_state = "hidden"
    workbook.active = workbook.sheetnames.index(import_sheet_name)

    base.apply_stable_openpyxl_sheet_formats(import_sheet, header_row=1)
    base.apply_stable_openpyxl_sheet_formats(query_sheet, header_row=1)
    base.save_openpyxl_workbook(workbook, output_path, overwrite=overwrite)


def export_payment_online(config: PaymentOnlineConfig) -> dict[str, Any]:
    source_data = extract_source_workbook_data(config.source_workbook_path, config.source_sheet_name)

    with tempfile.TemporaryDirectory(prefix="payment_online_") as temp_dir_text:
        temp_dir = Path(temp_dir_text)
        working_database_path = create_working_database_copy(config.access_database_path, temp_dir)
        import_rows_into_access(working_database_path, source_data.import_rows)

        query_headers, query_rows = fetch_query_rows(working_database_path, "Query2", order_by="[No]")
        round_updates = build_round_updates(query_headers, query_rows, source_data.round_label)
        apply_round_updates(working_database_path, round_updates)

        executed_queries = run_action_queries(working_database_path) if config.run_action_queries else []
        final_query_headers, final_query_rows = fetch_query_rows(working_database_path, "Query2", order_by="[No]")
        excel_text_values = build_excel_text_values(final_query_headers, final_query_rows)
        output_rows = build_output_rows(source_data.rows, excel_text_values)

        write_output_workbook(
            config.output_path,
            import_sheet_name=config.import_sheet_name,
            query_sheet_name=config.query_sheet_name,
            source_headers=source_data.headers,
            source_rows=output_rows,
            query_headers=final_query_headers,
            query_rows=final_query_rows,
            overwrite=config.overwrite,
        )

    return {
        "output_path": str(config.output_path),
        "source_workbook": str(config.source_workbook_path),
        "database_path": str(config.access_database_path),
        "row_count": len(source_data.rows),
        "query2_row_count": len(final_query_rows),
        "round_label": source_data.round_label,
        "run_action_queries": config.run_action_queries,
        "executed_queries": executed_queries,
        "action_deleted_existing_rows": sum(result.deleted_existing_rows for result in executed_queries),
        "action_inserted_rows": sum(result.inserted_rows for result in executed_queries),
        "action_skipped_duplicate_source_rows": sum(result.skipped_duplicate_source_rows for result in executed_queries),
    }


def main() -> int:
    args = parse_args()
    config = load_config(Path(args.config).resolve(), args)
    result = export_payment_online(config)
    print("Payment Online export completed successfully.")
    print(f"Source workbook: {result['source_workbook']}")
    print(f"Access database: {result['database_path']}")
    print(f"Round label: {result['round_label']}")
    print(f"Imported rows: {result['row_count']}")
    print(f"Query2 rows: {result['query2_row_count']}")
    print(f"Cleared existing action rows: {result['action_deleted_existing_rows']}")
    print(f"Inserted action rows: {result['action_inserted_rows']}")
    print(f"Skipped duplicate action source rows: {result['action_skipped_duplicate_source_rows']}")
    print(f"Executed action queries: {len(result['executed_queries'])}")
    for action_result in result["executed_queries"]:
        print(
            "  - "
            f"{action_result.query_name} -> {action_result.target_table} "
            f"(selected={action_result.selected_rows}, "
            f"cleared={action_result.deleted_existing_rows}, "
            f"inserted={action_result.inserted_rows}, "
            f"skipped_duplicates={action_result.skipped_duplicate_source_rows})"
        )
    print(f"Output workbook: {result['output_path']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
