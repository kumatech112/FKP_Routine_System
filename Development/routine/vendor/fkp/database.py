"""Linked SQL access with parameter binding, scoped backups and verified commits."""
from datetime import date, datetime
from decimal import Decimal
import re
import uuid

from .domain import Detail, Month, TABLES, ValidationError, key, number, validate_details
from .extract import close_measure
from .persistence import dump, now


class ComSql:
    def __init__(self, connection):
        import win32com.client
        self.dispatch = win32com.client.Dispatch
        self.connection = self.dispatch("ADODB.Connection")
        self.connection.ConnectionTimeout = 20
        self.connection.CommandTimeout = 120
        self.connection.Open("Provider=MSDASQL;" + connection.removeprefix("ODBC;"))
        self.connection.IsolationLevel = 1048576

    def run(self, sql, args=(), rows=True):
        command = self.dispatch("ADODB.Command")
        command.ActiveConnection = self.connection
        command.CommandText = sql
        command.CommandTimeout = 120
        for index, value in enumerate(args):
            text = str(value)
            parameter = command.CreateParameter(f"p{index}", 202, 1, max(1, len(text)), text)
            command.Parameters.Append(parameter)
        recordset = command.Execute()[0]
        result = []
        if rows and recordset is not None and recordset.State:
            names = [str(recordset.Fields(i).Name) for i in range(recordset.Fields.Count)]
            if not recordset.EOF:
                columns = recordset.GetRows()
                result = [dict(zip(names, row)) for row in zip(*columns)]
        if recordset is not None and recordset.State:
            recordset.Close()
        return result

    def begin(self):
        self.connection.BeginTrans()

    def commit(self):
        self.connection.CommitTrans()

    def rollback(self):
        self.connection.RollbackTrans()

    def close(self):
        self.connection.Close()


class Database:
    def __init__(self, path, sql=None):
        self.path = path
        self.sql = sql
        self.tables = {}

    def __enter__(self):
        if self.sql is not None:
            self.tables = {r: "[dbo].[" + t[0].removeprefix("dbo_") + "]" for r, t in TABLES.items()}
            return self
        import win32com.client
        dao = win32com.client.Dispatch("DAO.DBEngine.120")
        access = dao.OpenDatabase(str(self.path), False, True)
        connections = set()
        try:
            for report, (link, category) in TABLES.items():
                table = access.TableDefs(link)
                source = str(table.SourceTableName)
                if not re.fullmatch(r"dbo\.[A-Za-z0-9_]+", source):
                    raise ValidationError("ลิงก์ตารางฐานข้อมูลไม่ถูกต้อง")
                self.tables[report] = "[dbo].[" + source.split(".")[1] + "]"
                connection = str(table.Connect)
                if not connection.startswith("ODBC;"):
                    raise ValidationError("ตารางไม่ได้เชื่อมผ่าน ODBC ตามที่กำหนด")
                connections.add(connection)
        finally:
            access.Close()
        if len(connections) != 1:
            raise ValidationError("ตารางเชื่อมคนละฐานข้อมูล ไม่สามารถทำรายการร่วมกันได้")
        self.sql = ComSql(connections.pop())
        return self

    def __exit__(self, kind, error, traceback):
        self.sql.close()

    def schema(self):
        result = []
        for report, table in self.tables.items():
            fields = self.sql.run("SELECT c.name AS FieldName,t.name AS TypeName,c.max_length AS MaxLength,c.precision AS PrecisionValue,c.scale AS ScaleValue,c.is_nullable AS NullableValue,c.is_computed AS ComputedValue,c.default_object_id AS DefaultValue FROM sys.columns c JOIN sys.types t ON c.user_type_id=t.user_type_id WHERE c.object_id=OBJECT_ID(?)", [table])
            by_name = {key(r["FieldName"]): r for r in fields}
            category = TABLES[report][1]
            measures = ("Account", "OS", "Amount") if report in ("master", "billcode", "bucket") else ("Adjust Unit", "Adjust O/S", "Collected", "EFF %")
            required = [category, *measures, "Date_stamp", "Add_date", "Month_N", "NO"]
            if report in ("billcode", "bucket"):
                required += ["Entity", "Product", "ECM"]
            if report in ("daily", "adjust"):
                required += ["Product", "Bucket"]
            if any(key(name) not in by_name for name in required):
                raise ValidationError("โครงสร้างตารางไม่ครบ " + report)
            if any(by_name[key(name)]["ComputedValue"] for name in required):
                raise ValidationError("ช่องข้อมูลนำเข้าเป็นสูตรในฐาน " + report)
            for field in fields:
                if key(field["FieldName"]) not in {key(n) for n in required} and not field["NullableValue"] and not field["ComputedValue"] and not field["DefaultValue"]:
                    raise ValidationError("ฐานมีช่องบังคับที่ไม่ได้กำหนด " + str(field["FieldName"]))
            result.append({"report": report, "table": table, "fields": fields})
        return result

    @staticmethod
    def scope(row):
        if row.report in ("daily", "adjust"):
            return "[Date_stamp]=? AND [Product]=?", [row.asof, row.product]
        month = Month.parse(row.month)
        where, params = "[Date_stamp]>=? AND [Date_stamp]<?", [month.first.isoformat(), month.next.isoformat()]
        if row.report != "master":
            where += " AND [Entity]=? AND [Product]=? AND [ECM]=?"
            params += [row.entity, row.product, row.ecm]
        return where, params

    @staticmethod
    def payload(row):
        category = TABLES[row.report][1]
        daily = row.report in ("daily", "adjust")
        fields = [category, "Adjust Unit" if daily else "Account", "Adjust O/S" if daily else "OS", "Collected" if daily else "Amount", "Date_stamp", "Add_date", "Month_N", "NO"]
        values = [row.category, row.account, row.outstanding, row.amount, row.asof, now().date().isoformat(), row.month, str(uuid.uuid4())]
        if daily:
            fields += ["Product", "Bucket", "EFF %"]
            if row.efficiency is None:
                raise ValidationError("ไม่พบค่า EFF ของรายงานรายวัน")
            values += [row.product, row.bucket, row.efficiency]
        elif row.report != "master":
            fields += ["Entity", "Product", "ECM"]
            values += [row.entity, row.product, row.ecm]
        return fields, values

    @staticmethod
    def from_db(report, raw):
        value = {key(k): v for k, v in raw.items()}
        stamp = value["DATE_STAMP"]
        if isinstance(stamp, (date, datetime)):
            stamp = stamp.isoformat()[:10]
        else:
            stamp = str(stamp).strip()[:10]
        daily = report in ("daily", "adjust")
        return Detail(report, key(value["MONTH_N"]), stamp, key(value[key(TABLES[report][1])]), int(number(value["ADJUST UNIT" if daily else "ACCOUNT"], 0)), number(value["ADJUST O/S" if daily else "OS"], 4 if daily else 2), number(value["COLLECTED" if daily else "AMOUNT"], 4 if daily else 2), key(value.get("ENTITY", "ALL")), key(value.get("PRODUCT", "ALL")), key(value.get("ECM", "ALL")), key(value.get("BUCKET", "")) if daily else "", number(value["EFF %"]) if daily else None)

    def months(self, report):
        rows = self.sql.run(f"SELECT DISTINCT [Month_N] AS MonthValue FROM {self.tables[report]} WHERE [Month_N] IS NOT NULL")
        return [m.label for m in sorted({Month.parse(r["MonthValue"]) for r in rows})]

    def read(self, report, months):
        if not months:
            raise ValidationError("ยังไม่ได้เลือกเดือน")
        result = []
        for label in months:
            month = Month.parse(label)
            raw = self.sql.run(f"SELECT * FROM {self.tables[report]} WHERE [Date_stamp]>=? AND [Date_stamp]<? AND [Month_N]=?", [month.first.isoformat(), month.next.isoformat(), label])
            for row in raw:
                category = key(row[TABLES[report][1]])
                if category == "TOTAL" or category.startswith("GROUP"):
                    continue
                result.append(self.from_db(report, row))
        return validate_details(result)

    def import_rows(self, rows, replace, job):
        validate_details(rows)
        self.schema()
        scopes = {}
        for row in rows:
            where, args = self.scope(row)
            scopes[(row.report, where, tuple(args))] = row
        job.check()
        job.verify_sources()
        job.database_outcome = "unknown"
        self.sql.begin()
        committed = False
        try:
            old = []
            for report, where, args in scopes:
                values = self.sql.run(f"SELECT * FROM {self.tables[report]} WITH (UPDLOCK,HOLDLOCK) WHERE " + where, args)
                if values and not replace:
                    raise ValidationError("มีข้อมูลในเดือนหรือตัวกรองนี้แล้ว")
                old.append({"report": report, "where": where, "parameters": args, "rows": values})
            dump(job.evidence / ("before_" + uuid.uuid4().hex + ".json"), {"old": old, "incoming": [r.json() for r in rows]})
            job.check()
            if replace:
                for report, where, args in scopes:
                    self.sql.run(f"DELETE FROM {self.tables[report]} WHERE " + where, args, rows=False)
            for row in rows:
                fields, values = self.payload(row)
                sql = f"INSERT INTO {self.tables[row.report]} (" + ",".join("[" + f + "]" for f in fields) + ") VALUES (" + ",".join("?" for _ in fields) + ")"
                self.sql.run(sql, values, rows=False)
            self.sql.commit()
            committed = True
            job.database_outcome = "committed_unverified"
            dump(job.evidence / ("commit_" + uuid.uuid4().hex + ".json"), {"committed": True, "rows": len(rows), "scopes": [list(s) for s in scopes], "time": now().isoformat()})
            actual = []
            for report, where, args in scopes:
                actual.extend(self.from_db(report, r) for r in self.sql.run(f"SELECT * FROM {self.tables[report]} WHERE " + where, args))
            expected = {r.identity: r for r in rows}
            if len(actual) != len(expected) or {r.identity for r in actual} != set(expected):
                raise ValidationError("บันทึกแล้วแต่ตรวจสอบจำนวนไม่ผ่าน ห้ามเริ่มซ้ำก่อนตรวจฐาน")
            for row in actual:
                target = expected[row.identity]
                if row.asof != target.asof or (row.account, row.outstanding, row.amount) != (target.account, target.outstanding, target.amount) or row.efficiency != target.efficiency:
                    raise ValidationError("บันทึกแล้วแต่ตรวจสอบยอดไม่ผ่าน ห้ามเริ่มซ้ำก่อนตรวจฐาน")
            job.database_outcome = "verified"
            return len(actual)
        except Exception:
            if not committed:
                self.sql.rollback()
                job.database_outcome = "not_saved"
            raise
