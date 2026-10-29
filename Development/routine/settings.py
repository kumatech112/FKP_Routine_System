"""Portable multi-workflow settings. Secrets are protected for the current Windows user."""
from copy import deepcopy
from pathlib import Path
import base64
import json
import os
import sys

from .config import resource_root
from .runtime import bootstrap_windows
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import dump

TASKS = {"payment_online": "Payment Online", "performance": "Performance",
         "customer_documents": "Customer Documents", "enforcement": "Enforcement",
         "performance_success_rate": "Performance Success Rate"}


def default_settings():
    bootstrap_windows()
    seed = json.loads((resource_root() / "legacy_defaults.json").read_text(encoding="utf-8-sig"))
    seed["forward_mail"].setdefault("workflow_recipients", {}).setdefault("enforcement", {"to": [], "cc": []})
    flows = deepcopy(seed["workflows"])
    flows["enforcement"] = {
        "enabled": True, "label": "Enforcement", "store_name": seed["outlook"]["store_name"],
        "folder_path": ["กล่องจดหมายเข้า"], "network_folder": seed["files"]["network_folder"],
        "subject_regex": r"แจ้งผลขออนุมัติยึดอสังหาริมทรัพย์[_\s]*(?P<mail_date>\d{8})",
        "processor": "enforcement",
    }
    flows["performance_success_rate"] = {
        "enabled":False,"label":"Performance Success Rate","processor":"performance_success_rate",
        "store_name":seed["outlook"]["store_name"],"folder_path":["กล่องจดหมายเข้า","Performance"],
        "network_folder":r"\\10.0.0.243\pgp-data$","subject_regex":r".*Performance\s+success\s+rate.*",
    }
    seed["forward_mail"]["workflow_recipients"].setdefault("performance_success_rate",{"to":[],"cc":[]})
    for name, flow in flows.items():
        flow["extra_inbox"] = name in ("customer_documents","performance_success_rate")
    runtime = resource_root() / "runtime"
    return {
        "version": 2, "logs_root": r"\\10.0.0.97\Z Log Bot\FKP Routine System Log",
        "output_root": r"\\10.0.0.97\Z Collect Bot\FKP Routine System",
        "database": r"\\10.0.0.97\kuma$\X\FKP Performance.accdb",
        "payment_database": r"\\10.0.0.97\kuma$\X\FKP Payment Online Round.mdb",
        "workflows": flows, "outlook": deepcopy(seed["outlook"]),
        "mail": deepcopy(seed["forward_mail"]), "archive": {
            "archive_processed_files": seed["files"].get("archive_processed_files", True),
            "processed_archive_folder": seed["files"]["processed_archive_folder"],
        },
        "gpg_path": str(runtime / "gpg" / "gpg2.exe"),
        "gpg_home": "",
        "ocr_path": str(runtime / "ocr" / "tesseract.exe"),
        "poll": int(seed["app"].get("poll_interval_seconds", 25)),
        "scan_hours": int(seed["outlook"].get("received_within_hours",24)),
        "scan_limit": int(seed["outlook"].get("max_scan_items",200)),
        "secrets": {"passphrase": "", "zip_password": ""},
        "geometry": "1180x780+20+20",
    }


def settings_path():
    bootstrap_windows()
    override = os.environ.get("FKP_ROUTINE_PROFILE_ROOT")
    return Path(override or (Path(os.environ["LOCALAPPDATA"]) / "FKP_Routine_System")) / "settings.json"


def seal(data):
    import win32crypt
    return win32crypt.CryptProtectData(data, "FKP Routine System", None, None, None, 0)


def unseal(data):
    import win32crypt
    return win32crypt.CryptUnprotectData(data, None, None, None, 0)[1]


def validate(values):
    values["mail"].setdefault("workflow_recipients", {}).setdefault("enforcement", {"to": [], "cc": []})
    values["mail"]["workflow_recipients"].setdefault("performance_success_rate",{"to":[],"cc":[]})
    if set(values["workflows"]) != set(TASKS):
        raise ValidationError("รายการงานไม่ครบตามที่กำหนด")
    for field in ("logs_root", "output_root", "database", "payment_database", "gpg_path", "ocr_path"):
        if not Path(values[field]).is_absolute():
            raise ValidationError("ระบุพาธเต็ม: " + field)
    if values.get("gpg_home") and not Path(values["gpg_home"]).is_absolute():
        raise ValidationError("ระบุพาธเต็มของโฟลเดอร์กุญแจ PGP หรือเว้นว่าง")
    logs, collect = Path(values["logs_root"]).resolve(), Path(values["output_root"]).resolve()
    if logs == collect or logs in collect.parents or collect in logs.parents:
        raise ValidationError("โฟลเดอร์ Log/Process ต้องแยกจากงานสำเร็จ")
    if any("d:\\kuma bot" in str(values[f]).replace("/", "\\").casefold()
           for f in ("logs_root", "output_root")):
        raise ValidationError("ไม่ใช้ D:\\KUMA BOT เป็นพื้นที่ทำงาน")
    if not 1 <= int(values["poll"]) <= 3600:
        raise ValidationError("ช่วงตรวจเมลต้องอยู่ระหว่าง 1–3600 วินาที")
    if not 1<=int(values.get("scan_hours",24))<=8760 or not 1<=int(values.get("scan_limit",200))<=10000:
        raise ValidationError("ระบุช่วงย้อนหลัง 1–8760 ชั่วโมง และจำนวนเมล 1–10000 ต่อโฟลเดอร์")
    import re
    for name, flow in values["workflows"].items():
        if not flow.get("store_name") or not flow.get("folder_path") or not Path(flow["network_folder"]).is_absolute():
            raise ValidationError("ตรวจกล่องเมลและโฟลเดอร์ต้นทางของ " + TASKS[name])
        if name != "customer_documents":
            re.compile(flow["subject_regex"])
    for task, flow in values["workflows"].items():
        if task != "enforcement" and flow.get("enabled"):
            recipients = values["mail"].get("workflow_recipients", {}).get(task, {}).get("to", [])
            if not recipients or not str(values["mail"].get("send_from", "")).strip():
                raise ValidationError("ระบุผู้ส่งและผู้รับของ " + TASKS[task])
    return values


def load(path=None):
    values = default_settings()
    target = Path(path or settings_path())
    if not target.exists():
        # Migration is a one-time local read, not a dependency on another PC's D: drive.
        legacy = Path(r"D:\Documents\Routine Bot System\config.json")
        if legacy.exists():
            original = json.loads(legacy.read_text(encoding="utf-8-sig"))
            values["secrets"] = {"passphrase": original.get("decrypt", {}).get("passphrase", ""),
                                 "zip_password": original.get("forward_mail", {}).get("zip_password", "")}
        return validate(values)
    bundled_tools = {key: values[key] for key in ("gpg_path", "ocr_path")}
    saved = json.loads(target.read_text(encoding="utf-8-sig"))
    encrypted = saved.pop("protected_secrets", "")
    saved.pop("secrets", None)
    for field,old_field in (("scan_hours","received_within_hours"),("scan_limit","max_scan_items")):
        if field not in saved:
            values[field]=int(saved.get("outlook",{}).get(old_field,values[field]))
    saved_workflows=saved.pop("workflows",{})
    for name,flow in saved_workflows.items():
        values["workflows"].setdefault(name,{}).update(flow)
    values.update(saved)
    # Managed tools follow the running release; private key folders and custom tools stay intact.
    import re
    for key, current in bundled_tools.items():
        parts = Path(values[key]).parts
        expected = ("_internal", "runtime", *Path(current).parts[-2:])
        if (len(parts) >= 5 and re.fullmatch(r"FKP_Routine_System_\d+\.\d+\.\d+", parts[-5], re.I)
                and tuple(s.casefold() for s in parts[-4:]) == tuple(s.casefold() for s in expected)):
            values[key] = current
    if encrypted:
        try:
            values["secrets"] = json.loads(unseal(base64.b64decode(encrypted)))
        except Exception as error:
            raise ValidationError("โปรไฟล์รหัสผ่านเป็นของ Windows ผู้ใช้อื่น กรุณาตั้งค่าของผู้ใช้นี้") from error
    return validate(values)


def save(values, path=None):
    validate(values)
    stored = deepcopy(values)
    stored["protected_secrets"] = base64.b64encode(seal(
        json.dumps(stored.pop("secrets"), ensure_ascii=False).encode("utf-8"))).decode("ascii")
    dump(path or settings_path(), stored)


def task_config(values, task):
    name = TASKS[task]
    return {"logs": str(Path(values["logs_root"]) / name),
            "output": str(Path(values["output_root"]) / name),
            "database": values["database"], "payment_database": values["payment_database"]}


def runtime_settings(values, task):
    scan = deepcopy(values["outlook"])
    scan.update(deepcopy(values["workflows"][task]))
    if task == "customer_documents":
        scan["subject_regex"] = r".*"
        scan["sender_contains"] = []
    # Read/unread is not a work queue: durable file/mail ledgers prevent replay.
    scan["received_within_hours"]=int(values.get("scan_hours",scan.get("received_within_hours",24)))
    scan["max_scan_items"]=int(values.get("scan_limit",scan.get("max_scan_items",200)))
    scan["unread_only"] = False
    scan["flexible_trigger"] = task
    mail = deepcopy(values["mail"])
    recipients = mail.get("workflow_recipients", {}).get(task, {"to": [], "cc": []})
    mail["to"], mail["cc"] = list(recipients.get("to", [])), list(recipients.get("cc", []))
    if task in ("enforcement","performance_success_rate"):
        mail["enabled"] = True
        mail["zip_password_required"] = True
    mail["zip_password"] = values["secrets"].get("zip_password", "")
    return {"scan": scan, "mail": mail, "decrypt": {
                "enabled": True, "gpg_path": values["gpg_path"], "gpg_home": values["gpg_home"],
                "passphrase": values["secrets"].get("passphrase", ""),
                "archive_password": values["secrets"].get("zip_password", "")},
            "archive": deepcopy(values["archive"]), "poll": values["poll"]}


def dependency_checks(values, task, write_probe=False):
    """Read-only checks; never change ACLs, trust policies, registry or Office."""
    import winreg
    problems = []
    if values["mail"].get("zip_attachment") and values["mail"].get("zip_password_required") and not values["secrets"].get("zip_password"):
        problems.append("ยังไม่ได้ตั้งรหัส ZIP ของผู้ใช้นี้")
    if task == "enforcement":
        recipients = values["mail"].get("workflow_recipients", {}).get(task, {}).get("to", [])
        if not recipients or not str(values["mail"].get("send_from", "")).strip():
            problems.append("ระบุ Enforcement To และบัญชีผู้ส่งในตั้งค่าอีเมล")
        if not values["secrets"].get("zip_password") and not values["mail"].get("zip_password_required"):
            problems.append("ระบุรหัส ZIP สำหรับ Enforcement")
        if not values["mail"].get("zip_attachment", True):
            problems.append("Enforcement ต้องเปิดแนบเป็น ZIP")
    if task=="performance_success_rate":
        if not values["mail"].get("zip_attachment",True) or not values["secrets"].get("zip_password"):
            problems.append("Success Rate ต้องเปิด ZIP และใส่รหัสก่อนส่ง")
    required = ["Outlook.Application"]
    if task == "performance":
        required += ["Excel.Application", "DAO.DBEngine.120"]
    elif task == "payment_online":
        required += ["DAO.DBEngine.120"]
    for progid in required:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                pass
        except OSError:
            friendly = {"Outlook.Application": "Outlook แบบคลาสสิก", "Excel.Application": "Microsoft Excel", "DAO.DBEngine.120": "ส่วนประกอบ Access 64-bit"}
            problems.append("ไม่พบ " + friendly.get(progid, progid))
    if not Path(values["gpg_path"]).is_file():
        problems.append("ไม่พบโปรแกรมถอด PGP")
    if values.get("gpg_home") and not Path(values["gpg_home"]).is_dir():
        problems.append("ไม่พบโฟลเดอร์กุญแจ PGP ของผู้ใช้นี้")
    if task == "enforcement" and not Path(values["ocr_path"]).is_file():
        problems.append("ไม่พบ Tesseract OCR")
    if task == "payment_online" and not Path(values["payment_database"]).is_file():
        problems.append("ไม่พบ Access Payment")
    if task == "performance" and not Path(values["database"]).is_file():
        problems.append("ไม่พบ Access Daily")
    for field in ("logs_root", "output_root"):
        directory = Path(values[field])
        existing = next((p for p in [directory, *directory.parents] if p.exists()), None)
        if existing is None or not existing.is_dir() or not os.access(existing, os.R_OK):
            problems.append("เข้าโฟลเดอร์ไม่ได้: " + str(directory))
        elif write_probe:
            try:
                import tempfile
                with tempfile.NamedTemporaryFile(prefix=".fkp_permission_", dir=existing, delete=True) as probe:
                    probe.write(b"FKP permission probe")
                    probe.flush()
            except OSError:
                problems.append("ไม่มีสิทธิ์เขียนโฟลเดอร์: " + str(directory))
    return problems
