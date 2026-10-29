"""Independent work queues, shared-share locks and durable side-effect states."""
from pathlib import Path
import hashlib
import json
import re
import socket
import threading
import uuid

from .config import engine_settings
from .daily import DailyDatabase
from .delivery import package_files
from .enforcement import image_mail_text, prepare_enforcement, approve_review
from .files import archive_source, extract_payload, unpack_archive
from .matching import body_names, locate_body_file, payment_file, success_rate_file
from .outlook import Outlook
from .payment import prepare_payment, commit_actions
from .service import PerformanceService, RoutineJob, mail_key, safe_trace
from .settings import TASKS, task_config, runtime_settings, dependency_checks, validate
from .vendor.fkp.domain import Cancelled, ValidationError
from .vendor.fkp.office import apartment
from .vendor.fkp.persistence import digest, dump, lease, now, publish_outputs, read


class FileWorkflowService(PerformanceService):
    def __init__(self, values, task, stop=None, notify=None, outlook=None):
        self.values, self.task = values, task
        super().__init__(task_config(values, task), stop, notify, outlook=outlook,
                         legacy_loader=lambda _: runtime_settings(values, task))

    def poll(self):
        runtime = runtime_settings(self.values, self.task)
        outlook = self.outlook or Outlook()
        completed = []
        from .monitoring import scan_mail
        for reference in scan_mail(outlook,runtime["scan"],self.notify,self.task):
            if self.stop is not None and self.stop.is_set():
                break
            if reference.match.get('trigger_error'):
                self.notify({'kind':'status','step':TASKS[self.task]+': '+reference.match['trigger_error']})
                continue
            try:
                candidates=self.candidates(reference,runtime,outlook)
            except (OSError,ValidationError) as error:
                self.notify({'kind':'status','step':TASKS[self.task]+': '+str(error)})
                continue
            if self.task == "enforcement":
                if not candidates:
                    continue
                if any(source is None for source, _ in candidates):
                    from .intake_status import missing_source_note
                    missing=next(name for source,name in candidates if source is None)
                    self.notify({"kind":"status","step":"Enforcement: "+missing_source_note(runtime,reference.match,self.task,missing)})
                    continue
                logical_names = "\0".join(sorted(name.casefold() for _, name in candidates))
                unit_key = hashlib.sha256((mail_key(reference)+"\0"+logical_names).encode()).hexdigest()
                if unit_key in self.ledger.data["mails"]:
                    from .enforcement_flow import can_retry_before_mail
                    old_key=self.ledger.data["mails"][unit_key].get("file")
                    if not can_retry_before_mail(self.ledger.data["files"].get(old_key,{})):
                        continue
                receipt = self.process_unit(reference, [source for source, _ in candidates], runtime, outlook, unit_key)
                if receipt:
                    completed.append(receipt)
                continue
            for source, logical_name in candidates:
                if source is None:
                    from .intake_status import missing_source_note
                    self.notify({"kind":"status","step":TASKS[self.task]+": "+missing_source_note(runtime,reference.match,self.task,logical_name)})
                    continue
                unit_key = hashlib.sha256((mail_key(reference)+"\0"+logical_name.casefold()).encode()).hexdigest()
                if unit_key in self.ledger.data["mails"]:
                    continue
                receipt = self.process_unit(reference, source, runtime, outlook, unit_key)
                if receipt:
                    completed.append(receipt)
        return completed

    def candidates(self,reference,runtime,outlook):
        if self.task=='payment_online':
            return [(payment_file(runtime['scan'],reference.match),'payment')]
        if self.task=='performance_success_rate':
            return [(success_rate_file(runtime['scan'],reference.match,reference.body),'success_rate')]
        body=reference.body
        if self.task=='enforcement':
            intake=Path(self.settings['work'])/'Intake'/mail_key(reference)
            images=outlook.inline_images(reference,intake)
            if images:
                ocr_job=RoutineJob(intake,self.settings,self.stop,self.notify)
                cached=intake/'mail_image_ocr.json'
                if cached.exists():
                    item=read(cached);body+='\n'+item['text']+'\n'+item['compact']
                else:body+='\n'+image_mail_text(images,self.values['ocr_path'],ocr_job)
        names=reference.match.get('file_names') or body_names(body,self.task)
        if names:
            return [(locate_body_file(runtime['scan']['network_folder'],name),name) for name in names]
        from .triggers import infer_sources
        found=infer_sources(runtime['scan']['network_folder'],reference.match,self.task)
        return [(path,path.name) for path in found] or [(None,'รอชื่อไฟล์หรือไฟล์ที่ตรงเมล')]

    def process_unit(self, reference, source, runtime, outlook, unit_key):
        if self.task == "enforcement":
            from .enforcement_flow import process_sources
            sources = source if isinstance(source, (list, tuple)) else [source]
            return process_sources(self, reference, sources, runtime, unit_key, outlook=outlook)
        fingerprint = digest(source)
        if fingerprint in self.ledger.data["files"]:
            previous = self.ledger.data["files"][fingerprint]
            self.ledger.data["mails"][unit_key] = {"file": fingerprint, "status": previous["status"]}
            self.ledger.save()
            return None
        job = RoutineJob(Path(self.settings["work"]) / (now().strftime("%Y%m%d_%H%M%S")+"_"+uuid.uuid4().hex[:8]),
                         self.settings, self.stop, self.notify)
        self.ledger.data["files"][fingerprint] = {
            "status": "processing", "database": "not_requested", "mail": "not_requested",
            "collect": "not_started", "source": str(source), "job": str(job.folder),
            "task": self.task, "host": socket.gethostname(),
        }
        self.ledger.update(unit_key, fingerprint)
        forwarded = None
        try:
            if self.task == "payment_online":
                workbook = unpack_archive(source, job, runtime["decrypt"], runtime["scan"]["excel_regex"])
                round_number = int(reference.match["round"])
                if not 1 <= round_number <= 12:
                    raise ValidationError("รอบ Payment ต้องอยู่ระหว่าง 1–12")
                artifacts = prepare_payment(workbook, self.values["payment_database"], job,
                                             run_actions=round_number != 1)
                result_files = [artifacts.output]
                attachment = package_files(result_files, runtime["mail"], job, artifacts.output.stem)
                attachment_hash = digest(attachment)
                forwarded = outlook.prepare_forward(reference, attachment, runtime["mail"])
                artifacts.verify()
                job.check()
                self.ledger.update(unit_key, fingerprint, database="pending")
                outcome = commit_actions(artifacts, job)
                self.ledger.update(unit_key, fingerprint, database=outcome["status"])
                rows = artifacts.source_rows
            elif self.task in ("customer_documents","performance_success_rate"):
                if self.task=="performance_success_rate" and (not runtime["mail"].get("zip_attachment") or not runtime["mail"].get("zip_password")):
                    raise ValidationError("Success Rate ต้อง ZIP ใส่รหัสก่อนส่ง")
                extracted = extract_payload(source, job, runtime["decrypt"])
                result_files = sorted(p for p in extracted.rglob("*") if p.is_file())
                if not result_files:
                    raise ValidationError("ไฟล์ที่แตกแล้วไม่มีเอกสาร")
                from types import SimpleNamespace
                delivery=job.folder/"Delivery"
                delivery.mkdir()
                attachment = package_files(result_files, runtime["mail"], SimpleNamespace(folder=delivery), re.sub(r"\.(?:zip|rar)(?:\.(?:gpg|pgp))?$", "", Path(source).name, flags=re.I),
                                           relative_root=extracted)
                attachment_hash = digest(attachment)
                forwarded = outlook.prepare_forward(reference, attachment, runtime["mail"])
                rows = len(result_files)
            if digest(source) != fingerprint:
                raise ValidationError("ไฟล์ต้นทางเปลี่ยนระหว่างทำงาน")
            job.check()
            if digest(attachment) != attachment_hash:
                raise ValidationError("ไฟล์แนบเปลี่ยนก่อนส่งเมล")
            self.ledger.update(unit_key, fingerprint, mail="pending")
            outcome = outlook.deliver(forwarded, runtime["mail"])
            forwarded = None
            if outcome not in ("sent", "draft_opened"):
                raise ValidationError("ผลส่งเมลไม่ชัดเจน ห้ามเริ่มซ้ำ")
            self.ledger.update(unit_key, fingerprint, mail=outcome)
            job.progress(95, "ส่งเมลแล้ว" if outcome=="sent" else "เปิด Draft แล้ว")
            if digest(attachment) != attachment_hash:
                raise ValidationError("ไฟล์แนบเปลี่ยนหลังส่งเมล")
            # Keep raw document trees in Process; the verified ZIP is the completed deliverable.
            publish = [str(attachment)]
            if self.task == "payment_online" and attachment != result_files[0]:
                publish.insert(0, str(result_files[0]))
            self.ledger.update(unit_key, fingerprint, collect="pending")
            outputs = publish_outputs(publish, self.settings, job)
            self.ledger.update(unit_key, fingerprint, collect="verified", outputs=outputs)
            if outcome == "sent":
                archive_source(source, fingerprint, runtime["archive"])
                outlook.complete(reference, runtime["scan"])
            status = "success" if outcome=="sent" else "draft_opened"
            self.ledger.update(unit_key, fingerprint, status=status)
            receipt = job.finish(TASKS[self.task], status, "จัดทำรายงานและส่งเมลแล้ว" if outcome=="sent"
                                 else "จัดทำรายงานและเปิด Draft แล้ว", {"outputs": outputs, "records": rows})
        except BaseException as error:
            state = self.ledger.data["files"][fingerprint]
            if state["database"] == "pending":
                state["database"] = getattr(job, "database_outcome", "unknown")
            if state["mail"] == "pending":
                state["mail"] = "unknown"
            status = "cancelled" if isinstance(error, Cancelled) else "failed"
            self.ledger.update(unit_key, fingerprint, status=status)
            job.log(safe_trace(error, runtime))
            detail = str(error) if isinstance(error, (ValidationError, Cancelled)) else "ทำงานไม่สำเร็จ กรุณาตรวจบันทึกงาน"
            receipt = job.finish(TASKS[self.task], status, detail, dict(state))
        finally:
            if forwarded is not None:
                try:
                    forwarded.Close(1)
                except Exception:
                    pass
        self.notify({"kind": "complete", "receipt": receipt})
        return receipt


def run(values, stop, notify=None, once=False):
    validate(values)
    notify = notify or (lambda event: None)
    selected = [key for key in TASKS if values["workflows"][key].get("enabled", False)]
    if not selected:
        raise ValidationError("เลือกงานอย่างน้อยหนึ่งประเภท")
    services = {}
    with apartment():
        for task in selected:
            task_notify = lambda event, name=task: notify({**event, "task": name})
            try:
                problems=dependency_checks(values,task,write_probe=True)
                if problems:raise ValidationError('; '.join(problems))
                if task=='performance':
                    with DailyDatabase(values['database']) as database:database.schema()
                    service=PerformanceService(task_config(values,task),stop,task_notify,
                        legacy_loader=lambda _,name=task:runtime_settings(values,name))
                else:service=FileWorkflowService(values,task,stop,task_notify)
                services[task]=service
            except Exception as error:
                detail=str(error) if isinstance(error,ValidationError) else 'เชื่อมส่วนประกอบของงานไม่สำเร็จ'
                task_notify({'kind':'scan','phase':'error','time':now().isoformat(),
                    'store':values['workflows'][task]['store_name'],'folder':'/'.join(values['workflows'][task]['folder_path']),
                    'step':'งานยังไม่พร้อม: '+detail})
        if not services:raise ValidationError('ไม่มีงานที่พร้อมเริ่ม กรุณาดูเหตุผลในสถานะตรวจเมล')
        while not stop.is_set():
            for task, service in services.items():
                if stop.is_set():
                    break
                lock_folder = Path(service.settings["logs"]) / "Logs" / "Locks"
                try:
                    with lease(lock_folder, "performance" if task == "performance" else "workflow"):
                        # Re-read the shared ledger while holding the cross-machine lease.
                        from .service import Ledger
                        service.ledger = Ledger(service.settings["logs"])
                        service.poll()
                except OSError as error:
                    notify({"kind":"scan","task":task,"phase":"error","time":now().isoformat(),
                        "store":values["workflows"][task]["store_name"],"folder":"/".join(values["workflows"][task]["folder_path"]),
                        "step":"เข้าไฟล์หรือโฟลเดอร์ไม่ได้ " + str(getattr(error,"filename",""))})
                except ValidationError as error:
                    notify({"kind":"scan","task":task,"phase":"error","time":now().isoformat(),
                        "store":values["workflows"][task]["store_name"],"folder":"/".join(values["workflows"][task]["folder_path"]),
                        "step":str(error)})
            if once:
                break
            from datetime import timedelta
            next_check = (now()+timedelta(seconds=int(values["poll"]))).isoformat()
            for task in selected:
                notify({"kind":"poll_wait","task":task,"time":now().isoformat(),"next_check":next_check})
            notify({"kind": "status", "step": "กำลังเฝ้าตรวจเมลงานที่เลือก"})
            stop.wait(int(values["poll"]))


def finalize_enforcement(review_path, edits, notify=None, stop=None, values=None, outlook=None):
    review_path = Path(review_path)
    context = read(review_path.parent / "review_context.json")
    if context.get("workflow_version",0) >= 4:
        from .enforcement_flow import finalize
        return finalize(review_path, edits, notify, stop, values=values, outlook=outlook)
    settings = context["settings"]
    with lease(Path(settings["logs"]) / "Logs" / "Locks", "workflow"):
        from .service import Ledger
        ledger = Ledger(settings["logs"])
        state = ledger.data["files"][context["source_sha256"]]
        if state["status"] != "needs_review":
            raise ValidationError("สถานะงานไม่ใช่รอตรวจทาน")
        job = RoutineJob(review_path.parent, settings, stop=stop, notify=notify)
        additional = []
        if context.get("delivery_required"):
            if state.get("mail") not in ("sent", "draft_opened"):
                raise ValidationError("ผลส่ง ZIP Enforcement ยังไม่ชัดเจน")
            attachment = Path(context["mail_attachment"])
            if not attachment.is_file() or digest(attachment) != context["mail_attachment_sha256"]:
                raise ValidationError("ZIP Enforcement ไม่ตรงกับไฟล์ที่ส่งเมล")
            additional.append(str(attachment))
        with lease(review_path.parent / "ReviewLocks", "edit"):
            outputs = approve_review(review_path, edits, job, settings, additional_outputs=additional)
        draft = context.get("delivery_required") and state.get("mail")=="draft_opened"
        status = "draft_opened" if draft else "success"
        ledger.update(context["unit_key"], context["source_sha256"], status=status,
                      collect="verified", outputs=outputs)
        if not draft:
            archive_source(context["source"], context["source_sha256"], context["runtime_archive"])
        receipt = job.finish("Enforcement", status, "ตรวจทานครบและเก็บไฟล์แล้ว",
                             {"outputs": outputs, "mail": state.get("mail", "not_requested")})
        if notify:
            notify({"kind": "complete", "receipt": receipt})
        return receipt


def save_review_cell(review_path, cell_id, value):
    review_path = Path(review_path)
    with lease(review_path.parent / "ReviewLocks", "edit"):
        review = read(review_path)
        if review["status"] != "needs_review":
            raise ValidationError("งานนี้ไม่อยู่ในสถานะรอตรวจทานแล้ว")
        cell = next((c for c in review["cells"] if c["id"] == cell_id), None)
        if cell is None:
            raise ValidationError("ไม่พบช่องที่เลือก")
        cell.update(value=value, state="reviewed")
        dump(review_path, review)
        return review
