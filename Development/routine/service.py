"""Durable Performance orchestration with separate SQL/mail delivery outcomes."""
from pathlib import Path
import hashlib
import json
import traceback
import uuid

from .config import engine_settings, load_legacy
from .daily import DailyDatabase, process_daily
from .files import archive_source, find_archive, make_attachment, mail_date, unpack_archive
from .outlook import Outlook
from .vendor.fkp.domain import Cancelled, ValidationError
from .vendor.fkp.office import apartment
from .vendor.fkp.persistence import Job, digest, dump, lease, now, publish_outputs, read


class RoutineJob(Job):
    def __init__(self, folder, settings, stop=None, notify=None):
        super().__init__(folder, settings)
        self.stop = stop
        self.notify = notify or (lambda value: None)

    def check(self):
        super().check()
        if self.stop is not None and self.stop.is_set():
            raise Cancelled("หยุดการทำงานแล้ว")

    def event(self, value):
        super().event(value)
        self.notify(value)


class Ledger:
    def __init__(self, logs):
        self.path = Path(logs) / "Logs" / "state.json"
        self.data = read(self.path) if self.path.exists() else {"mails": {}, "files": {}}
        if set(self.data) != {"mails", "files"}:
            raise ValidationError("สถานะงานเสียหาย ต้องตรวจสอบก่อนเริ่ม")

    def save(self):
        dump(self.path, self.data)

    def update(self, mail_key, file_key, **changes):
        current = self.data["files"][file_key]
        current.update(changes, updated=now().isoformat())
        self.data["mails"][mail_key] = {"file": file_key, "status": current["status"]}
        self.save()


def mail_key(reference):
    return hashlib.sha256((reference.store_id + "\0" + reference.entry_id).encode("utf-8")).hexdigest()


def safe_trace(error, legacy):
    text = "".join(traceback.format_exception(error))
    for section, field in (("decrypt", "passphrase"), ("decrypt", "archive_password"), ("mail", "zip_password")):
        secret = str(legacy.get(section, {}).get(field, ""))
        if secret:
            text = text.replace(secret, "<redacted>")
    import re
    return re.sub(r"(?i)(PWD|PASSWORD|PASSPHRASE)\s*=\s*[^;\r\n]*", r"\1=<redacted>", text)


class MailSourceMismatch(ValidationError):
    """A valid file belongs to another mail cutoff; it remains available to the right mail."""


class PerformanceService:
    def __init__(self, config, stop=None, notify=None, outlook=None,
                 database_factory=DailyDatabase, processor=process_daily, legacy_loader=load_legacy):
        self.config = dict(config)
        self.settings = engine_settings(config)
        self.stop = stop
        self.notify = notify or (lambda value: None)
        self.outlook = outlook
        self.database_factory = database_factory
        self.processor = processor
        self.legacy_loader = legacy_loader
        self.ledger = Ledger(self.settings["logs"])

    def poll(self):
        legacy = self.legacy_loader(self.config)
        if not legacy["scan"].get("enabled", True):
            self.notify({"kind": "status", "step": "Performance ถูกปิดในการตั้งค่าเดิม"})
            return []
        outlook = self.outlook or Outlook()
        completed = []
        from .monitoring import scan_mail
        for reference in scan_mail(outlook,legacy["scan"],self.notify,"performance",workflow=False):
            if self.stop is not None and self.stop.is_set():
                break
            key = mail_key(reference)
            if self.ledger.data["mails"].get(key,{}).get("file"):
                continue
            if reference.match.get("trigger_error"):
                self.notify({"kind":"status","step":"Performance: "+reference.match["trigger_error"]})
                continue
            try:
                source = find_archive(legacy["scan"], reference.match, reference.body)
                if source is None:
                    self.notify({"kind": "status", "step": "รอไฟล์ Performance ที่ตรงเมล"})
                    continue
                rejected=self.ledger.data["mails"].get(key,{})
                if rejected.get("rejected_source_sha256")==digest(source):
                    continue
                result = self.process_message(reference, source, legacy, outlook)
                if result is not None:
                    completed.append(result)
            except Exception as error:
                # Matching failures have no SQL/mail effects. Record once for this mail.
                self.ledger.data["mails"][key] = {"status": "needs_review", "error": type(error).__name__}
                self.ledger.save()
                self.notify({"kind": "status", "step": str(error) if isinstance(error, ValidationError)
                             else "ตรวจเมลหรือไฟล์ไม่สำเร็จ กรุณาตรวจบันทึกงาน"})
        return completed

    def process_message(self, reference, source, legacy, outlook):
        mkey, fkey = mail_key(reference), digest(source)
        if fkey in self.ledger.data["files"]:
            previous = self.ledger.data["files"][fkey]
            self.ledger.data["mails"][mkey] = {"file": fkey, "status": previous["status"]}
            self.ledger.save()
            return None
        folder = Path(self.settings["work"]) / (now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8])
        job = RoutineJob(folder, self.settings, self.stop, self.notify)
        self.ledger.data["files"][fkey] = {
            "status": "processing", "database": "not_started", "mail": "not_started",
            "collect": "not_started", "source": str(source), "job": str(job.folder),
        }
        self.ledger.update(mkey, fkey)
        forwarded = None
        try:
            source_file = unpack_archive(source, job, legacy["decrypt"], legacy["scan"]["excel_regex"])
            if digest(source) != fkey:
                raise ValidationError("ไฟล์เครือข่ายเปลี่ยนระหว่างคัดลอก")
            artifacts = self.processor(source_file, self.settings, job)
            if not reference.match.get("mail_date"):
                asofs={row.asof for row in artifacts.records}
                if len(asofs)!=1:
                    raise ValidationError("ข้อมูล Daily ต้องมี as of เดียวที่ชัดเจน")
                from datetime import datetime
                reference.match["mail_date"]=datetime.strptime(next(iter(asofs)),"%Y-%m-%d").strftime("%d%m%Y")
            if any(row.asof != mail_date(reference.match).isoformat() for row in artifacts.records):
                raise MailSourceMismatch("วันที่ข้อมูลในขั้น 05 ไม่ตรงกับ as of ของเมล")
            artifacts.verify()
            attachment = make_attachment(artifacts.mail_file, legacy["mail"], job)
            attachment_hash = digest(attachment)
            # Resolve sender/recipients before committing financial data.
            forwarded = outlook.prepare_forward(reference, attachment, legacy["mail"])
            job.check()
            artifacts.verify()
            self.ledger.update(mkey, fkey, database="pending")
            with self.database_factory(self.settings["database"]) as database:
                database.import_rows(list(artifacts.records), False, job)
            self.ledger.update(mkey, fkey, database="verified")
            job.progress(90, "ตรวจสอบฐาน Daily สำเร็จ")
            job.check()
            artifacts.verify()
            if digest(attachment) != attachment_hash:
                raise ValidationError("ไฟล์แนบขั้น 01 เปลี่ยนก่อนส่งเมล")
            self.ledger.update(mkey, fkey, mail="pending")
            outcome = outlook.deliver(forwarded, legacy["mail"])
            forwarded = None
            if outcome not in ("sent", "draft_opened"):
                raise ValidationError("ไม่ทราบผลการส่งเมล ต้องตรวจสอบก่อนเริ่มซ้ำ")
            self.ledger.update(mkey, fkey, mail=outcome)
            job.progress(95, "ส่งเมลแล้ว" if outcome == "sent" else "เปิด Draft แล้ว")
            files = [str(artifacts.mail_file), str(artifacts.database_file)]
            if attachment not in (artifacts.mail_file, artifacts.database_file):
                files.append(str(attachment))
            self.ledger.update(mkey, fkey, collect="pending")
            outputs = publish_outputs(files, self.settings, job)
            self.ledger.update(mkey, fkey, collect="verified", outputs=outputs)
            archived = None
            if outcome == "sent":
                self.ledger.update(mkey, fkey, archive="pending")
                archived = archive_source(source, fkey, legacy["archive"])
                self.ledger.update(mkey, fkey, archive="archived" if archived else "not_requested")
                outlook.complete(reference, legacy["scan"])
            status = "success" if outcome == "sent" else "draft_opened"
            self.ledger.update(mkey, fkey, status=status)
            result = {"mode": "daily", "mail_stage": "01", "database_stage": "05",
                      "database": "verified", "mail": outcome, "records": len(artifacts.records),
                      "outputs": outputs, "archived_source": archived}
            receipt = job.finish("Performance Daily", status, "ลงฐาน Daily และส่งเมลเสร็จแล้ว"
                                 if outcome == "sent" else "ลงฐาน Daily แล้ว และเปิด Draft", result)
            self.notify({"kind": "complete", "receipt": receipt})
            return receipt
        except BaseException as error:
            current = self.ledger.data["files"][fkey]
            if current["database"] == "pending":
                current["database"] = getattr(job, "database_outcome", "unknown")
            if current["mail"] == "pending":
                current["mail"] = "unknown"
            status = "cancelled" if isinstance(error, Cancelled) else "failed"
            self.ledger.update(mkey, fkey, status=status)
            job.log(safe_trace(error, legacy))
            detail = str(error) if isinstance(error, (ValidationError, Cancelled)) else "ทำงานไม่สำเร็จ กรุณาตรวจสอบบันทึกงาน"
            receipt = job.finish("Performance Daily", status, detail, dict(current))
            if isinstance(error,MailSourceMismatch):
                del self.ledger.data["files"][fkey]
                self.ledger.data["mails"][mkey]={"status":"source_mismatch","rejected_source_sha256":fkey}
                self.ledger.save()
            self.notify({"kind": "complete", "receipt": receipt})
            return receipt
        finally:
            if forwarded is not None:
                try:
                    forwarded.Close(1)
                except Exception:
                    pass


def monitor(config, stop, notify=None, once=False):
    settings = engine_settings(config)
    notify = notify or (lambda event: None)
    with lease(Path(settings["logs"]) / "Logs" / "Locks", "performance"):
        with apartment():
            try:
                with DailyDatabase(settings["database"]) as database:
                    database.schema()
            except Exception as error:
                raise ValidationError("เชื่อมฐาน Daily ไม่สำเร็จ กรุณาตรวจ ODBC/OLE DB และสิทธิ์เข้าถึง") from error
            service = PerformanceService(config, stop, notify)
            while not stop.is_set():
                service.poll()
                if once:
                    break
                notify({"kind": "status", "step": "รอเมล Performance"})
                stop.wait(load_legacy(config)["poll"])
