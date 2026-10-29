"""Enforcement batches: decrypt, table-only OCR, password ZIP and automatic mail."""
from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace
import hashlib
import re
import uuid

from .delivery import package_files
from .enforcement import prepare_enforcement,prepare_review
from .files import extract_payload,archive_source
from .outlook import Outlook,MailRef
from .service import RoutineJob,Ledger,safe_trace
from .settings import load,runtime_settings
from .vendor.fkp.domain import ValidationError,Cancelled
from .vendor.fkp.office import apartment
from .vendor.fkp.persistence import digest,dump,read,lease,now,copy_new,publish_outputs


PRE_MAIL_FIX="encrypted-zip-1"


def can_retry_before_mail(state):
    return (state.get("status")=="failed" and state.get("mail") in ("not_started","not_requested")
        and state.get("collect","not_started")=="not_started" and not state.get("review")
        and state.get("pre_mail_fix")!=PRE_MAIL_FIX)


def copy_pdf(source,root,relative):
    target=Path(root)/relative
    original=target
    number=2
    while target.exists():
        target=original.with_name(f'{original.stem} ({number}){original.suffix}')
        number+=1
    return copy_new(source,target)


def process_sources(service,reference,sources,runtime,unit_key,outlook=None):
    sources=[Path(s) for s in sources]
    manifest=[{"path":str(s),"sha256":digest(s)} for s in sources]
    hashes=sorted(m["sha256"] for m in manifest)
    fingerprint=hashes[0] if len(hashes)==1 else hashlib.sha256("\0".join(hashes).encode()).hexdigest()
    previous_attempts=[]
    for old_key,old in service.ledger.data["files"].items():
        old_hashes=set(old.get("source_hashes",[old_key]))
        if old_key==fingerprint and can_retry_before_mail(old):
            previous_attempts=list(old.get("previous_attempts",[]))+[{k:old.get(k) for k in ("status","mail","job","error_phase","updated")} ]
            continue
        if old_key==fingerprint or old_hashes.intersection(hashes):
            service.ledger.data["mails"][unit_key]={"file":old_key,"status":old["status"]}
            service.ledger.save()
            service.notify({"kind":"status","step":"Enforcement: พบไฟล์ที่เคยรับงานแล้ว ตรวจประวัติก่อนทำซ้ำ"})
            return None
    job=RoutineJob(Path(service.settings["work"])/(now().strftime("%Y%m%d_%H%M%S")+"_"+uuid.uuid4().hex[:8]),
                   service.settings,service.stop,service.notify)
    service.ledger.data["files"][fingerprint]={"status":"processing","database":"not_requested",
        "mail":"not_started","collect":"not_started","source_hashes":hashes,"sources":manifest,"job":str(job.folder),
        "task":"enforcement","pre_mail_fix":PRE_MAIL_FIX,"previous_attempts":previous_attempts}
    service.ledger.update(unit_key,fingerprint)
    phase="extract"
    try:
        job.progress(5,"กำลังถอดและแตกไฟล์ Enforcement")
        pdf_root=job.folder/"PDFs"
        pdf_root.mkdir()
        for index,source in enumerate(sources,1):
            folder=job.folder/"Inputs"/str(index)
            folder.mkdir(parents=True)
            extracted=extract_payload(source,SimpleNamespace(folder=folder,check=job.check),
                                      runtime["decrypt"],allow_pdf=True)
            pdfs=sorted(p for p in extracted.rglob("*") if p.is_file() and p.suffix.casefold()==".pdf")
            if not pdfs:
                raise ValidationError("ไฟล์ที่ถอดได้ไม่มี PDF")
            for pdf in pdfs:
                copy_pdf(pdf,pdf_root,pdf.relative_to(extracted))
        phase="ocr"
        review=prepare_enforcement(pdf_root,job,service.values["ocr_path"])
        for item in manifest:
            if digest(item["path"])!=item["sha256"]:
                raise ValidationError("ต้นทาง Enforcement เปลี่ยนระหว่าง OCR")
        pdf_manifest=[{"path":str(p),"sha256":digest(p)} for p in sorted(pdf_root.rglob("*.pdf"))]
        # Private passwords stay in the current Windows profile, never in job context.
        context={"workflow_version":5,"task":"enforcement","source_sha256":fingerprint,
            "sources":manifest,"pdfs":pdf_manifest,"pdf_root":str(pdf_root),"unit_key":unit_key,
            "settings":service.settings,"runtime_archive":runtime["archive"],
            "reference":{"entry_id":reference.entry_id,"store_id":reference.store_id,"subject":reference.subject}}
        dump(job.folder/"review_context.json",context)
        review.update(status="ocr_complete",review_required=False,human_reviewed=False)
        dump(job.folder/"enforcement_review.json",review)
        service.ledger.update(unit_key,fingerprint,status="ready_to_send",review=str(job.folder/"enforcement_review.json"))
        job.progress(85,"OCR เสร็จแล้ว กำลัง ZIP และส่งเมล")
        phase="delivery"
        receipt=_deliver(job.folder/"enforcement_review.json",{},service.notify,service.stop,
            values=service.values,outlook=outlook or service.outlook,ledger=service.ledger,
            job=job,require_review=False,record_failure=False)
    except BaseException as error:
        status="cancelled" if isinstance(error,Cancelled) else "failed"
        if service.ledger.data["files"][fingerprint]["status"] != "ready_to_send":
            service.ledger.update(unit_key,fingerprint,status=status,error_phase=phase)
        job.log(safe_trace(error,runtime))
        detail=str(error) if isinstance(error,(ValidationError,Cancelled)) else {"extract":"เตรียมไฟล์ Enforcement ไม่สำเร็จ กรุณาตรวจบันทึก","ocr":"OCR ไม่สำเร็จ กรุณาตรวจบันทึก","delivery":"จัดส่ง Enforcement ไม่สำเร็จ กรุณาตรวจบันทึก"}[phase]
        receipt=job.finish("Enforcement",status,detail,dict(service.ledger.data["files"][fingerprint]))
    service.notify({"kind":"complete","receipt":receipt})
    return receipt


def finalize(review_path,edits,notify=None,stop=None,values=None,outlook=None,require_review=False):
    review_path=Path(review_path)
    context=read(review_path.parent/"review_context.json")
    settings=context["settings"]
    with lease(Path(settings["logs"])/"Logs"/"Locks","workflow"):
        receipt=_deliver(review_path,edits,notify,stop,values,outlook,require_review=require_review)
    if notify:notify({"kind":"complete","receipt":receipt})
    return receipt


def _deliver(review_path,edits,notify=None,stop=None,values=None,outlook=None,ledger=None,job=None,require_review=False,record_failure=True):
    review_path=Path(review_path)
    context=read(review_path.parent/"review_context.json")
    settings=context["settings"]
    ledger=ledger or Ledger(settings["logs"])
    key=context["source_sha256"]
    state=ledger.data["files"][key]
    if state["status"] not in ("ready_to_send","needs_review") or state.get("mail") not in ("not_started","not_requested"):
        raise ValidationError("สถานะงานไม่พร้อมส่ง ห้ามส่งซ้ำอัตโนมัติ")
    job=job or RoutineJob(review_path.parent,settings,stop,notify)
    runtime=runtime_settings(values if values is not None else load(),"enforcement")
    mail=runtime["mail"]
    forwarded=None
    try:
        if not mail.get("to") or not str(mail.get("send_from","")).strip():
            raise ValidationError("ระบุ Enforcement To และบัญชีผู้ส่งในตั้งค่าอีเมล")
        if not mail.get("zip_attachment") or not mail.get("zip_password"):
            raise ValidationError("Enforcement ต้อง ZIP ใส่รหัสก่อนส่ง")
        with lease(review_path.parent/"ReviewLocks","edit"):
            review=prepare_review(review_path,edits,require_review=require_review)
            review.update(review_required=require_review,human_reviewed=bool(require_review))
            for item in [*context["sources"],*context["pdfs"]]:
                if not Path(item["path"]).is_file() or digest(item["path"])!=item["sha256"]:
                    raise ValidationError("PDF หรือต้นทางเปลี่ยนหลัง OCR กรุณาตรวจสอบก่อนส่ง")
            payload=job.folder/"MailPayload"/uuid.uuid4().hex[:8]
            payload.mkdir(parents=True)
            files=[]
            for item in context["pdfs"]:
                source=Path(item["path"])
                relative=source.relative_to(context["pdf_root"])
                if context.get("workflow_version",4)<=4 and re.fullmatch(r"Input_\d+",relative.parts[0]):
                    relative=Path(*relative.parts[1:])
                target=copy_pdf(source,payload/"PDF",relative)
                files.append(target)
            excel=Path(review["output"])
            target=payload/excel.name
            copy_new(excel,target)
            files.append(target)
            delivery=job.folder/"Delivery"/uuid.uuid4().hex[:8]
            delivery.mkdir(parents=True)
            attachment=package_files(files,{**mail,"zip_password_required":True},
                SimpleNamespace(folder=delivery),"Enforcement_PDF_Excel",relative_root=payload)
            attachment_hash=digest(attachment)
            excel_hash=digest(excel)
            ref=context["reference"]
            reference=MailRef(ref["entry_id"],ref["store_id"],ref["subject"],"",{})
            # Outlook objects belong to this worker's COM apartment.
            with apartment() if outlook is None else nullcontext():
                adapter=outlook or Outlook()
                forwarded=adapter.prepare_forward(reference,attachment,mail)
                job.check()
                if digest(attachment)!=attachment_hash:
                    raise ValidationError("ZIP เปลี่ยนก่อนส่งเมล")
                ledger.update(context["unit_key"],key,mail="pending",mail_attachment=str(attachment),
                    mail_attachment_sha256=attachment_hash,recipients={"to":mail["to"],"cc":mail.get("cc",[])})
                outcome=adapter.deliver(forwarded,mail)
                forwarded=None
                if outcome not in ("sent","draft_opened"):
                    raise ValidationError("ผลส่งเมลไม่ชัดเจน ห้ามส่งซ้ำ")
                ledger.update(context["unit_key"],key,mail=outcome)
            if digest(attachment)!=attachment_hash or digest(excel)!=excel_hash:
                raise ValidationError("ไฟล์เปลี่ยนหลังส่งเมล")
            outputs=publish_outputs([str(excel),str(attachment)],settings,job)
            review.update(status="success",outputs=outputs)
            dump(review_path,review)
        ledger.update(context["unit_key"],key,collect="verified",outputs=outputs)
        archived=[]
        if outcome=="sent":
            for item in context["sources"]:
                archived.append(archive_source(item["path"],item["sha256"],context["runtime_archive"]))
        status="success" if outcome=="sent" else "draft_opened"
        ledger.update(context["unit_key"],key,status=status)
        receipt=job.finish("Enforcement",status,"ส่ง PDF และ Excel ใน ZIP แล้ว" if outcome=="sent"
            else "เปิด Draft ZIP ที่มี PDF และ Excel แล้ว",
            {"outputs":outputs,"mail":outcome,"pdfs":len(context["pdfs"]),"archived_sources":archived})

        return receipt
    except BaseException as error:
        if state.get("mail")=="pending":
            state["mail"]="unknown"
        status="ready_to_send" if state.get("mail") in ("not_started","not_requested") else "failed"
        ledger.update(context["unit_key"],key,status=status)
        job.log(safe_trace(error,runtime))
        if record_failure:
            job.finish("Enforcement","failed",str(error) if isinstance(error,ValidationError) else "ส่งงานไม่สำเร็จ กรุณาตรวจบันทึก",
                       dict(state))
        raise
    finally:
        if forwarded is not None:
            try:forwarded.Close(1)
            except Exception:pass
