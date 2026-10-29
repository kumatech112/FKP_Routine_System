"""Sequential Enforcement preparation followed by one verified batch delivery."""
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

PRE_MAIL_FIX="sequential-batch-1"

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

def workbook_name(source,index,total):
    if total==1:return "Enforcement.xlsx"
    stem=re.sub(r'\.(?:gpg|pgp)$','',Path(source).name,flags=re.I)
    stem=re.sub(r'\.(?:pdf|zip|rar)$','',stem,flags=re.I)
    stem=re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',stem).strip(' .')[:150]
    return (stem or f'Enforcement_{index}')+'.xlsx'

def process_sources(service,reference,sources,runtime,unit_key,outlook=None):
    # Collapse repeated references to the same path, retaining distinct source files.
    sources=list(dict.fromkeys(Path(s).resolve() for s in sources))
    if not sources:raise ValidationError("ไม่พบไฟล์ Enforcement สำหรับจัดชุด")
    manifest=[{"path":str(s),"sha256":digest(s)} for s in sources]
    hashes=sorted(m["sha256"] for m in manifest)
    fingerprint=hashes[0] if len(hashes)==1 else hashlib.sha256("\0".join(hashes).encode()).hexdigest()
    previous_attempts=[]
    for old_key,old in service.ledger.data["files"].items():
        old_hashes=set(old.get("source_hashes",[old_key]))
        if old_key==fingerprint and can_retry_before_mail(old):
            previous_attempts=list(old.get("previous_attempts",[]))+[{k:old.get(k) for k in ("status","mail","job","error_phase","updated")}]
            continue
        if old_key==fingerprint or old_hashes.intersection(hashes):
            service.ledger.data["mails"][unit_key]={"file":old_key,"status":old["status"]}
            service.ledger.save()
            detail="OCR เสร็จแล้ว รอส่งงานที่ค้าง" if old.get("status") in ("ready_to_send","needs_review") else "พบไฟล์ที่เคยรับงานแล้ว ตรวจประวัติก่อนทำซ้ำ"
            service.notify({"kind":"status","step":"Enforcement: "+detail})
            return None
    job=RoutineJob(Path(service.settings["work"])/(now().strftime("%Y%m%d_%H%M%S")+"_"+uuid.uuid4().hex[:8]),
                   service.settings,service.stop,service.notify)
    service.ledger.data["files"][fingerprint]={"status":"processing","database":"not_requested",
        "mail":"not_started","collect":"not_started","source_hashes":hashes,"sources":manifest,"job":str(job.folder),
        "task":"enforcement","pre_mail_fix":PRE_MAIL_FIX,"previous_attempts":previous_attempts,"sets":[]}
    service.ledger.update(unit_key,fingerprint)
    phase="extract"
    try:
        pdf_root=job.folder/"PDFs";pdf_root.mkdir()
        sets=[]
        used_names=set()
        total=len(sources)
        for index,source in enumerate(sources,1):
            job.check()
            phase="extract"
            job.progress(5+75*(index-1)/total,f"ถอดไฟล์ชุด {index}/{total}: {source.name}")
            folder=job.folder/"Inputs"/str(index);folder.mkdir(parents=True)
            extracted=extract_payload(source,SimpleNamespace(folder=folder,check=job.check),
                                      runtime["decrypt"],allow_pdf=True)
            pdfs=sorted(p for p in extracted.rglob("*") if p.is_file() and p.suffix.casefold()==".pdf")
            if not pdfs:raise ValidationError(f"ไฟล์ชุด {index} ที่ถอดได้ไม่มี PDF")
            staged=[]
            for pdf in pdfs:
                copied=copy_pdf(pdf,pdf_root,Path(pdf.name))
                staged.append({"path":str(copied),"sha256":digest(copied)})
            phase="ocr"
            set_folder=job.folder/"Sets"/str(index);set_folder.mkdir(parents=True)
            def progress(percent,message,number=index):
                job.progress(5+75*(number-1+min(100,max(0,percent))/100)/total,
                             f"ชุด {number}/{total}: "+message)
            set_job=SimpleNamespace(folder=set_folder,check=job.check,log=job.log,progress=progress)
            job.progress(5+75*(index-1)/total,f"OCR ชุด {index}/{total}: {source.name}")
            review=prepare_enforcement(extracted,set_job,service.values["ocr_path"])
            output=Path(review["output"])
            target=output.with_name(workbook_name(source,index,total))
            original=target;number=2
            while target.name.casefold() in used_names:
                target=original.with_name(f"{original.stem} ({number}){original.suffix}");number+=1
            used_names.add(target.name.casefold())
            if output!=target:
                output.rename(target);review["output"]=str(target)
            review.update(status="ocr_complete",review_required=False,human_reviewed=False)
            review_path=set_folder/"enforcement_review.json";dump(review_path,review)
            if digest(source)!=manifest[index-1]["sha256"]:
                raise ValidationError("ต้นทาง Enforcement เปลี่ยนระหว่าง OCR")
            item={"index":index,"source":manifest[index-1],"review":str(review_path),
                  "output":review["output"],"page_count":review["page_count"],
                  "cell_count":review["cell_count"],"pdfs":staged}
            sets.append(item)
            service.ledger.update(unit_key,fingerprint,sets=list(sets),prepared_sets=index,total_sets=total)
            job.progress(5+75*index/total,f"เตรียม PDF และ Excel แล้ว {index}/{total} ชุด")
        for item in manifest:
            if digest(item["path"])!=item["sha256"]:raise ValidationError("ต้นทาง Enforcement เปลี่ยนระหว่างทำงาน")
        pdf_manifest=[item for prepared in sets for item in prepared["pdfs"]]
        context={"workflow_version":6,"task":"enforcement","source_sha256":fingerprint,
            "sources":manifest,"pdfs":pdf_manifest,"pdf_root":str(pdf_root),"unit_key":unit_key,
            "settings":service.settings,"runtime_archive":runtime["archive"],
            "reference":{"entry_id":reference.entry_id,"store_id":reference.store_id,"subject":reference.subject}}
        dump(job.folder/"review_context.json",context)
        batch={"version":4,"status":"ocr_complete","sets":sets,
               "page_count":sum(s["page_count"] for s in sets),"cell_count":sum(s["cell_count"] for s in sets),
               "output":sets[0]["output"],"review_required":False,"human_reviewed":False}
        review_path=job.folder/"enforcement_review.json";dump(review_path,batch)
        service.ledger.update(unit_key,fingerprint,status="ready_to_send",review=str(review_path))
        job.progress(85,f"OCR ครบ {total} ชุด กำลังรวมไฟล์เพื่อส่งเมลครั้งเดียว")
        phase="delivery"
        receipt=_deliver(review_path,{},service.notify,service.stop,values=service.values,
            outlook=outlook or service.outlook,ledger=service.ledger,job=job,require_review=False,record_failure=False)
    except BaseException as error:
        status="cancelled" if isinstance(error,Cancelled) else "failed"
        if service.ledger.data["files"][fingerprint]["status"]!="ready_to_send":
            service.ledger.update(unit_key,fingerprint,status=status,error_phase=phase)
        job.log(safe_trace(error,runtime))
        detail=str(error) if isinstance(error,(ValidationError,Cancelled)) else {
            "extract":"เตรียมไฟล์ Enforcement ไม่สำเร็จ กรุณาตรวจบันทึก",
            "ocr":"OCR ไม่สำเร็จ กรุณาตรวจบันทึก","delivery":"จัดส่ง Enforcement ไม่สำเร็จ กรุณาตรวจบันทึก"}[phase]
        receipt=job.finish("Enforcement",status,detail,dict(service.ledger.data["files"][fingerprint]))
    service.notify({"kind":"complete","receipt":receipt})
    return receipt

def finalize(review_path,edits,notify=None,stop=None,values=None,outlook=None,require_review=False):
    review_path=Path(review_path);context=read(review_path.parent/"review_context.json")
    with lease(Path(context["settings"]["logs"])/"Logs"/"Locks","workflow"):
        receipt=_deliver(review_path,edits,notify,stop,values,outlook,require_review=require_review)
    if notify:notify({"kind":"complete","receipt":receipt})
    return receipt

def prepare_workbooks(review_path,edits,require_review):
    batch=read(review_path)
    if not batch.get("sets"):
        review=prepare_review(review_path,edits,require_review=require_review)
        return review,[Path(review["output"])]
    if batch["status"]=="success":raise ValidationError("งานนี้ยืนยันสำเร็จแล้ว")
    paths=[]
    for item in batch["sets"]:
        review=prepare_review(item["review"],{},require_review=require_review)
        path=Path(review["output"])
        if path!=Path(item["output"]):raise ValidationError("ตำแหน่ง Excel ของชุดงานไม่ตรงกัน")
        paths.append(path)
    return batch,paths

def _deliver(review_path,edits,notify=None,stop=None,values=None,outlook=None,ledger=None,job=None,require_review=False,record_failure=True):
    review_path=Path(review_path);context=read(review_path.parent/"review_context.json")
    settings=context["settings"];ledger=ledger or Ledger(settings["logs"]);key=context["source_sha256"]
    state=ledger.data["files"][key]
    if state["status"] not in ("ready_to_send","needs_review") or state.get("mail") not in ("not_started","not_requested"):
        raise ValidationError("สถานะงานไม่พร้อมส่ง ห้ามส่งซ้ำอัตโนมัติ")
    job=job or RoutineJob(review_path.parent,settings,stop,notify)
    runtime=runtime_settings(values if values is not None else load(),"enforcement");mail=runtime["mail"]
    forwarded=None
    try:
        if not mail.get("to") or not str(mail.get("send_from","")).strip():
            raise ValidationError("OCR ครบแล้ว รอระบุ Enforcement To และบัญชีผู้ส่งในตั้งค่าอีเมล แล้วส่งงาน OCR ที่ค้าง")
        if not mail.get("zip_attachment") or not mail.get("zip_password"):
            raise ValidationError("Enforcement ต้อง ZIP ใส่รหัสก่อนส่ง")
        with lease(review_path.parent/"ReviewLocks","edit"):
            review,excels=prepare_workbooks(review_path,edits,require_review)
            review.update(review_required=require_review,human_reviewed=bool(require_review))
            for item in [*context["sources"],*context["pdfs"]]:
                if not Path(item["path"]).is_file() or digest(item["path"])!=item["sha256"]:
                    raise ValidationError("PDF หรือต้นทางเปลี่ยนหลัง OCR กรุณาตรวจสอบก่อนส่ง")
            payload=job.folder/"MailPayload"/uuid.uuid4().hex[:8];payload.mkdir(parents=True)
            files=[]
            # One flat delivery folder; source subfolders and Input_1 are private implementation paths.
            for item in context["pdfs"]:
                source=Path(item["path"])
                files.append(copy_pdf(source,payload,Path(source.name)))
            excel_hashes={str(path):digest(path) for path in excels}
            for excel in excels:files.append(copy_pdf(excel,payload,Path(excel.name)))
            delivery=job.folder/"Delivery"/uuid.uuid4().hex[:8];delivery.mkdir(parents=True)
            attachment=package_files(files,{**mail,"zip_password_required":True},
                SimpleNamespace(folder=delivery),"Enforcement_PDF_Excel",relative_root=payload)
            attachment_hash=digest(attachment);ref=context["reference"]
            reference=MailRef(ref["entry_id"],ref["store_id"],ref["subject"],"",{})
            with apartment() if outlook is None else nullcontext():
                adapter=outlook or Outlook();forwarded=adapter.prepare_forward(reference,attachment,mail)
                job.check()
                if digest(attachment)!=attachment_hash:raise ValidationError("ZIP เปลี่ยนก่อนส่งเมล")
                ledger.update(context["unit_key"],key,mail="pending",mail_attachment=str(attachment),
                    mail_attachment_sha256=attachment_hash,recipients={"to":mail["to"],"cc":mail.get("cc",[])})
                outcome=adapter.deliver(forwarded,mail);forwarded=None
                if outcome not in ("sent","draft_opened"):raise ValidationError("ผลส่งเมลไม่ชัดเจน ห้ามส่งซ้ำ")
                ledger.update(context["unit_key"],key,mail=outcome)
            if digest(attachment)!=attachment_hash or any(digest(path)!=expected for path,expected in excel_hashes.items()):
                raise ValidationError("ไฟล์เปลี่ยนหลังส่งเมล")
            outputs=publish_outputs([*[str(path) for path in excels],str(attachment)],settings,job)
            review.update(status="success",outputs=outputs);dump(review_path,review)
        ledger.update(context["unit_key"],key,collect="verified",outputs=outputs)
        archived=[]
        if outcome=="sent":
            for item in context["sources"]:archived.append(archive_source(item["path"],item["sha256"],context["runtime_archive"]))
        status="success" if outcome=="sent" else "draft_opened"
        ledger.update(context["unit_key"],key,status=status)
        count=len(review.get("sets",[])) or 1
        receipt=job.finish("Enforcement",status,
            f"รวม PDF และ Excel {count} ชุดใน ZIP และส่งเมลครั้งเดียวแล้ว" if outcome=="sent"
            else f"เปิด Draft ZIP รวม PDF และ Excel {count} ชุดแล้ว",
            {"outputs":outputs,"mail":outcome,"pdfs":len(context["pdfs"]),"sets":count,"excel_files":len(excels),"archived_sources":archived})
        return receipt
    except BaseException as error:
        if state.get("mail")=="pending":state["mail"]="unknown"
        status="ready_to_send" if state.get("mail") in ("not_started","not_requested") else "failed"
        ledger.update(context["unit_key"],key,status=status);job.log(safe_trace(error,runtime))
        if record_failure:
            job.finish("Enforcement","failed",str(error) if isinstance(error,ValidationError)
                else "ส่งงานไม่สำเร็จ กรุณาตรวจบันทึก",dict(state))
        raise
    finally:
        if forwarded is not None:
            try:forwarded.Close(1)
            except Exception:pass
