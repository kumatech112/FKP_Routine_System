"""Payment Online adapter: original mapping/queries, staged MDB, transactional Action writes."""
from dataclasses import dataclass
from pathlib import Path
import re

from .vendor import legacy_payment as legacy
from .vendor.fkp.database import ComSql
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import copy_new, digest, dump
from .settings import seal


@dataclass(frozen=True)
class ActionBatch:
    query: str
    table: str
    fields: tuple
    rows: tuple
    connection: str


@dataclass(frozen=True)
class PaymentArtifacts:
    output: Path
    batches: tuple
    source_rows: int
    output_hash: str

    def verify(self):
        if digest(self.output) != self.output_hash:
            raise ValidationError("ไฟล์ Payment เปลี่ยนหลังตรวจสอบ")


def prepare_payment(source, database_path, job, run_actions=True):
    staged_source = job.stage(source)
    # Preserve the original filename for round extraction, not a temporary UUID.
    data = legacy.extract_source_workbook_data(staged_source, None)
    if not data.rows:
        raise ValidationError("ไฟล์ Payment ไม่มีรายการ")
    original_hash = digest(database_path)
    working = copy_new(database_path, job.folder / "Payment_work.mdb")
    try:
        legacy.import_rows_into_access(working, data.import_rows)
        headers, query_rows = legacy.fetch_query_rows(working, "Query2", order_by="[No]")
        if len(query_rows) != len(data.rows):
            raise ValidationError("จำนวนรายการ Query2 ไม่ตรงกับต้นทาง ห้ามจับคู่ตามแถว")
        number_index = headers.index("No")
        ids = [row[number_index] for row in query_rows]
        if len(set(ids)) != len(ids):
            raise ValidationError("Query2 มีเลข No ซ้ำ ต้องตรวจข้อมูลเชื่อมก่อน")
        legacy.apply_round_updates(working, legacy.build_round_updates(headers, query_rows, data.round_label))
        headers, query_rows = legacy.fetch_query_rows(working, "Query2", order_by="[No]")
        values = legacy.build_output_rows(data.rows, legacy.build_excel_text_values(headers, query_rows))
        output_stem=re.sub(r"_processed", "", Path(source).stem, flags=re.I)
        output = job.folder / (output_stem + ".xlsx")
        legacy.write_output_workbook(output, import_sheet_name="Table1", query_sheet_name="Query1",
            source_headers=data.headers, source_rows=values, query_headers=headers, query_rows=query_rows,
            overwrite=False)
        batches = []
        engine, database = legacy.base.open_dao_database(working)
        try:
            if run_actions:
                for query in database.QueryDefs:
                    name = str(query.Name)
                    if "Payment Online" not in name or "Action" not in name:
                        continue
                    if int(query.Type) != 64:
                        raise ValidationError("Action Payment ต้องเป็น Append Query")
                    plan = legacy.parse_action_insert_sql(str(query.SQL))
                    rows, duplicates = legacy.unique_action_rows(legacy.fetch_action_insert_rows(database, plan))
                    table = database.TableDefs(plan.target_table)
                    target, connection = str(table.SourceTableName), str(table.Connect)
                    if not re.fullmatch(r"dbo\.[\w ]+", target) or not connection.startswith("ODBC;"):
                        raise ValidationError("ปลายทาง Action ไม่ใช่ลิงก์ SQL ที่รองรับ")
                    if any(any(row.get(field) is None for field in plan.target_columns) for row in rows):
                        raise ValidationError("Action มีข้อมูลบังคับว่าง")
                    batches.append(ActionBatch(name, "[dbo].[" + target.split(".", 1)[1] + "]",
                        tuple(plan.target_columns), tuple(rows), connection))
        finally:
            database.Close()
        if digest(database_path) != original_hash:
            raise ValidationError("Access ต้นฉบับเปลี่ยนระหว่างทำงาน")
        dump(job.evidence / "payment_prepared.json", {
            "source_rows": len(data.rows), "query_rows": len(query_rows), "round": data.round_label,
            "actions": [{"query": b.query, "rows": len(b.rows)} for b in batches],
            "database_source_hash": original_hash, "output": str(output),
        })
        job.progress(80, "เตรียมรายงาน Payment แล้ว")
        return PaymentArtifacts(output, tuple(batches), len(data.rows), digest(output))
    finally:
        # A staged MDB can contain saved linked-table credentials: retain an encrypted
        # recovery copy for this Windows user, never ship/store a plaintext MDB in logs.
        if working.exists():
            encrypted = working.with_suffix(".mdb.dpapi")
            with encrypted.open("xb") as stream:
                stream.write(seal(working.read_bytes()))
            working.unlink()


def sql_scope(row):
    clauses, values = [], []
    for field in legacy.ACTION_DEDUPLICATION_KEY_COLUMNS:
        value = row[field]
        if value is None:
            clauses.append("[" + field + "] IS NULL")
        else:
            clauses.append("[" + field + "]=?")
            values.append(value)
    return " AND ".join(clauses), values


def commit_actions(artifacts, job, sql_factory=ComSql):
    artifacts.verify()
    active = [batch for batch in artifacts.batches if batch.rows]
    if not active:
        return {"status": "not_required", "rows": 0}
    groups = {}
    for batch in active:
        groups.setdefault(batch.connection, []).append(batch)
    opened = {}
    outcomes = []
    total = 0
    job.database_outcome = "unknown"
    try:
        # Validate every destination before the first write. Connections retain the
        # original Access links/credentials; no cross-database target is invented.
        for index, (connection, batches) in enumerate(groups.items(), 1):
            sql = opened[connection] = sql_factory(connection)
            for batch in batches:
                fields = ",".join("[" + f.replace("]", "]]") + "]" for f in batch.fields)
                sql.run(f"SELECT TOP (0) {fields} FROM {batch.table}")
        for group_number, (connection, batches) in enumerate(groups.items(), 1):
            job.check()
            sql = opened[connection]
            committed, commit_started = False, False
            state = {"group": group_number, "queries": [b.query for b in batches], "status": "pending"}
            outcomes.append(state)
            dump(job.evidence / "payment_groups.json", outcomes)
            sql.begin()
            try:
                before = []
                for batch in batches:
                    for row in batch.rows:
                        where, args = sql_scope(row)
                        old = sql.run(f"SELECT * FROM {batch.table} WITH (UPDLOCK,HOLDLOCK) WHERE "+where, args)
                        before.append({"table": batch.table, "scope": args, "rows": old})
                dump(job.evidence / f"payment_before_{group_number}.json", {"before": before})
                job.check()
                written = 0
                for batch in batches:
                    columns = ",".join("[" + f.replace("]", "]]") + "]" for f in batch.fields)
                    for row in batch.rows:
                        where, args = sql_scope(row)
                        sql.run(f"DELETE FROM {batch.table} WHERE "+where, args, rows=False)
                        sql.run(f"INSERT INTO {batch.table} ({columns}) VALUES ("+
                            ",".join("?" for _ in batch.fields)+")",
                            [row[f] for f in batch.fields], rows=False)
                        written += 1
                state["status"] = "commit_pending"
                dump(job.evidence / "payment_groups.json", outcomes)
                commit_started = True
                sql.commit()
                committed = True
                state["status"] = "committed_unverified"
                dump(job.evidence / "payment_groups.json", outcomes)
                for batch in batches:
                    for row in batch.rows:
                        where, args = sql_scope(row)
                        actual = sql.run(f"SELECT * FROM {batch.table} WHERE "+where, args)
                        if len(actual) != 1 or any(str(actual[0].get(f, "")).strip()!=str(row[f]).strip()
                                                  for f in batch.fields):
                            raise ValidationError("Action บันทึกแล้วแต่ตรวจเทียบไม่ผ่าน ห้ามเริ่มซ้ำ")
                total += written
                state.update(status="verified", rows=written)
                job.database_outcome = "partial_verified"
                dump(job.evidence / "payment_groups.json", outcomes)
            except BaseException:
                if not committed:
                    try:
                        sql.rollback()
                        state["status"] = "unknown" if commit_started else "not_saved"
                    except Exception:
                        state["status"] = "unknown"
                job.database_outcome = "partial_verified" if any(s["status"]=="verified" for s in outcomes) else state["status"]
                dump(job.evidence / "payment_groups.json", outcomes)
                raise
        job.database_outcome = "verified"
        return {"status": "verified", "rows": total, "groups": outcomes}
    finally:
        for sql in opened.values():
            try:
                sql.close()
            except Exception:
                pass
