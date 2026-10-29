"""Routine settings; legacy mail/decryption secrets are read, never copied."""
from copy import deepcopy
from pathlib import Path
import json
import re
import sys

from .vendor.fkp.domain import PRODUCTS, ValidationError
from .vendor.fkp.persistence import defaults as engine_defaults

DAILY_PRODUCTS = tuple(p for p in PRODUCTS if p != "BAYKAMC")


def resource_root():
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))


def default_config():
    return {
        "legacy_config": r"D:\Documents\Routine Bot System\config.json",
        "logs": r"\\10.0.0.97\Z Log Bot\FKP Routine System Log\Performance",
        "output": r"\\10.0.0.97\Z Collect Bot\FKP Routine System\Performance",
        "database": r"\\10.0.0.97\kuma$\X\FKP Performance.accdb",
        "geometry": "1180x780+20+20",
    }


def config_path():
    return (Path(sys.executable).parent if getattr(sys, "frozen", False)
            else Path(__file__).resolve().parents[2] / "Program") / "settings.local.json"


def load_config(path=None):
    values = default_config()
    target = Path(path or config_path())
    if target.exists():
        saved = json.loads(target.read_text(encoding="utf-8-sig"))
        if set(saved) - set(values):
            raise ValidationError("การตั้งค่าโปรแกรมมีช่องที่ไม่รองรับ")
        values.update(saved)
    for name in ("legacy_config", "logs", "output", "database"):
        if not Path(values[name]).is_absolute():
            raise ValidationError("ระบุพาธเต็มของ " + name)
    logs, output = Path(values["logs"]).resolve(), Path(values["output"]).resolve()
    if logs == output or logs in output.parents or output in logs.parents:
        raise ValidationError("แยกโฟลเดอร์งานสำเร็จออกจาก Log/Process")
    for name in ("logs", "output"):
        if str(values[name]).replace("/", "\\").casefold().startswith("d:\\kuma bot"):
            raise ValidationError("ไม่ใช้โฟลเดอร์ D:\\KUMA BOT ในโปรแกรมใหม่")
    return values


def load_legacy(config):
    path = Path(config["legacy_config"])
    values = json.loads(path.read_text(encoding="utf-8-sig"))
    flow = deepcopy(values["workflows"]["performance"])
    if flow.get("processor") != "performance":
        raise ValidationError("ไม่พบการตั้งค่าเมล Performance เดิม")
    scan = deepcopy(values["outlook"])
    scan.update(flow)
    for field in ("subject_regex", "network_file_regex", "excel_regex"):
        re.compile(scan[field])
    mail = deepcopy(values["forward_mail"])
    destination = mail.get("workflow_recipients", {}).get("performance")
    if destination is not None:
        mail["to"] = list(destination.get("to", []))
        mail["cc"] = list(destination.get("cc", []))
    else:
        mail["to"] = list(mail.get("recipients", [])) + list(mail.get("groups", []))
        mail["cc"] = list(mail.get("cc_recipients", [])) + list(mail.get("cc_groups", []))
    if not mail.get("enabled") or not mail.get("attach_payment_output"):
        raise ValidationError("งาน Performance ต้องเปิดส่งเมลและแนบผลลัพธ์ตามที่กำหนด")
    if not mail["to"] or not mail.get("send_from"):
        raise ValidationError("ตรวจสอบผู้ส่งและผู้รับเมล Performance")
    decrypt = deepcopy(values.get("decrypt", {}))
    gpg = Path(decrypt.get("gpg_path", "gpg/gpg2.exe"))
    if not gpg.is_absolute():
        gpg = path.parent / gpg
    decrypt["gpg_path"] = str(gpg)
    return {
        "scan": scan, "mail": mail, "decrypt": decrypt,
        "archive": deepcopy(values.get("files", {})),
        "poll": max(1, int(values.get("app", {}).get("poll_interval_seconds", 25))),
    }


def engine_settings(config):
    values = engine_defaults(resource_root())
    for key in ("logs", "output", "database"):
        values[key] = config[key]
    values["work"] = str(Path(config["logs"]) / "Process")
    values["templates"] = {
        "daily": str(resource_root() / "Templates" / "Daily_Recovery_Performance.xls"),
        "index": str(resource_root() / "Templates" / "index Bucket.xls"),
    }
    return values
