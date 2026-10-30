"""Atomic local state, shared-job evidence, locks and source preservation."""
from contextlib import contextmanager
from datetime import datetime, timezone, timedelta
from pathlib import Path
import hashlib
import json
import os
import shutil
import tempfile
import time
import uuid

from .domain import Cancelled, ValidationError, TEMPLATES

BANGKOK = timezone(timedelta(hours=7))


def now():
    return datetime.now(BANGKOK)


def dump(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2, default=str)
            stream.flush()
            os.fsync(stream.fileno())
        replace_file(temporary, path)
    finally:
        if temporary.exists():
            try:
                temporary.unlink()
            except OSError:
                pass


def replace_file(temporary, target):
    for attempt in range(3):
        try:
            os.replace(temporary, target)
            return
        except PermissionError:
            if attempt == 2:
                raise
            time.sleep(.05 * (attempt + 1))


def read(path):
    if os.name != "nt":
        return json.loads(Path(path).read_text(encoding="utf-8-sig"))
    import win32file
    handle = win32file.CreateFile(str(path), win32file.GENERIC_READ, win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE, None, win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
    try:
        size = win32file.GetFileSize(handle)
        content = win32file.ReadFile(handle, size)[1] if size else b""
        return json.loads(content.decode("utf-8-sig"))
    finally:
        handle.Close()


def read_events(path, offset=0):
    if os.name == "nt":
        import win32file
        handle = win32file.CreateFile(str(path), win32file.GENERIC_READ, win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE, None, win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
        try:
            size = win32file.GetFileSize(handle)
            win32file.SetFilePointer(handle, offset, 0)
            content = win32file.ReadFile(handle, size - offset)[1] if size > offset else b""
        finally:
            handle.Close()
    else:
        with Path(path).open("rb") as stream:
            stream.seek(offset)
            content = stream.read()
    complete = content.rfind(b"\n") + 1
    events = [json.loads(line.decode("utf-8")) for line in content[:complete].splitlines() if line]
    return offset + complete, events


def digest(path):
    result = hashlib.sha256()
    with Path(path).open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            result.update(chunk)
    return result.hexdigest()


def copy_new(source, target):
    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    with Path(source).open("rb") as incoming, target.open("xb") as outgoing:
        shutil.copyfileobj(incoming, outgoing)
    if digest(source) != digest(target):
        raise ValidationError("ตรวจสอบสำเนาไฟล์ไม่ผ่าน")
    return target


def profiles():
    return [{"sheet": s, "entity": "TESCO" if s in ("LTS", "TES") else s, "product": s if s in ("LTS", "TES") else "ALL", "product_group": "TCS" if s in ("LTS", "TES") else "ALL", "enabled": s != "TESCO"} for s in ("ALL", "KRUNGSRI", "GCS", "AYCAP", "TESCO", "LTS", "TES", "BAY", "KAMC")]


def defaults(root):
    return {"database": r"\\10.0.0.97\kuma$\X\FKP Performance.accdb", "logs": r"\\10.0.0.97\Z Log Bot\FKP Performance Log", "work": r"\\10.0.0.97\Z Log Bot\FKP Performance Log\Process", "output": r"\\10.0.0.97\Z Collect Bot\FKP Performance", "templates": {k: str(Path(root) / "Templates" / v) for k, v in TEMPLATES.items()}, "oa_code": "1015-TRADE", "profiles": profiles(), "geometry": "1180x780"}


def log_path(settings, *parts):
    return Path(settings["logs"]) / "Logs" / Path(*parts)


def process_path(settings, *parts):
    return Path(settings["logs"]) / "Process" / Path(*parts)


def ensure_layout(settings):
    for name in ("Logs", "Process", "Archive"):
        (Path(settings["logs"]) / name).mkdir(parents=True, exist_ok=True)


def publish_outputs(paths, settings, job, category=None):
    """Expose only complete verified files, after optional database readback."""
    if not paths:
        return []
    job.check()
    job.verify_sources()
    output = Path(settings["output"])
    output.parent.mkdir(parents=True, exist_ok=True)
    output.mkdir(exist_ok=True)
    if category not in (None, "Daily", "Adjust", "Master"):
        raise ValueError("Unknown completed report category")
    destination = (output / category if category else output) / job.id
    if destination.exists():
        raise ValidationError("มีโฟลเดอร์ผลลัพธ์นี้แล้ว กรุณาตรวจสอบก่อนเริ่มใหม่")
    # A short-lived delivery buffer on the destination share allows atomic rename
    # across the separate Log/Collect SMB shares. Workbooks stay in Process.
    temporary = Path(tempfile.mkdtemp(prefix=".FKP_Performance_Delivery_", dir=output.parent))
    try:
        for source in paths:
            job.check()
            copy_new(source, temporary / Path(source).name)
        job.check()
        destination.parent.mkdir(parents=True, exist_ok=True)
        os.rename(temporary, destination)
    finally:
        if temporary.exists():
            resolved = temporary.resolve()
            if resolved.parent != output.parent.resolve() or not resolved.name.startswith(".FKP_Performance_Delivery_"):
                raise RuntimeError("Unsafe delivery cleanup path")
            shutil.rmtree(resolved)
    published = [str(destination / Path(source).name) for source in paths]
    dump(job.evidence / "delivery.json", {"files": published, "verified": True})
    return published


def validate_settings(settings):
    for field in ("database", "logs", "work", "output"):
        if not str(settings.get(field, "")).strip() or not Path(settings[field]).is_absolute():
            raise ValidationError("ระบุพาธเต็มของฐานข้อมูลและโฟลเดอร์")
    expected_work = process_path(settings)
    if Path(settings["work"]) != expected_work:
        raise ValidationError("โฟลเดอร์ประมวลผลต้องอยู่ใน Process ใต้โฟลเดอร์บันทึกงาน")
    output = str(Path(settings["output"])).casefold().rstrip("\\/")
    logs = str(Path(settings["logs"])).casefold().rstrip("\\/")
    if output == logs or output.startswith(logs + os.sep) or logs.startswith(output + os.sep):
        raise ValidationError("โฟลเดอร์ไฟล์ที่เสร็จแล้วต้องแยกจากบันทึกงาน")
    if not str(settings.get("oa_code", "")).strip():
        raise ValidationError("ระบุ OA Code")
    names = set()
    for profile in settings["profiles"]:
        if profile["sheet"] in names or not all(str(profile[k]).strip() for k in ("sheet", "entity", "product", "product_group")):
            raise ValidationError("ตัวกรองชีตว่างหรือชื่อชีตซ้ำ")
        names.add(profile["sheet"])
    return settings


@contextmanager
def lease(folder, name):
    from ...locking import lease as resource_lease
    with resource_lease(folder, name):
        yield


class Job:
    def __init__(self, folder, settings):
        ensure_layout(settings)
        self.folder = Path(folder)
        self.settings = settings
        self.id = self.folder.name + "_" + uuid.uuid4().hex[:8]
        self.folder.mkdir(parents=True, exist_ok=True)
        self.evidence = log_path(settings, "Jobs", self.id)
        self.evidence.mkdir(parents=True, exist_ok=True)
        self.started = time.monotonic()
        self.percent = 0
        self.sources = []

    def check(self):
        if (self.folder / "cancel").exists():
            raise Cancelled("ยกเลิกงานแล้ว")

    def progress(self, percent, step):
        self.percent = max(self.percent, min(99, int(percent)))
        self.event({"kind": "progress", "percent": self.percent, "step": step})

    def event(self, value):
        content = (json.dumps(value, ensure_ascii=False) + "\n").encode("utf-8")
        with (self.folder / "events.jsonl").open("ab", buffering=0) as stream:
            stream.write(content)

    def stage(self, source):
        self.check()
        source = Path(source)
        before = digest(source)
        target = copy_new(source, self.folder / (uuid.uuid4().hex[:10] + "_" + source.name))
        if digest(source) != before:
            raise ValidationError("ไฟล์ต้นฉบับเปลี่ยนระหว่างอ่านข้อมูล")
        self.sources.append((str(source), before))
        dump(self.evidence / "sources.json", self.sources)
        from .source_safety import disable_refresh
        disable_refresh(target)
        return target

    def verify_sources(self):
        if any(digest(path) != expected for path, expected in self.sources):
            raise ValidationError("ไฟล์ต้นฉบับเปลี่ยน กรุณาอ่านข้อมูลใหม่")

    def log(self, text):
        with (self.evidence / "technical.log").open("a", encoding="utf-8") as stream:
            stream.write(now().isoformat() + " " + text + "\n")

    def queue(self, states):
        dump(self.evidence / "queue.json", states)
        self.event({"kind": "queue", "rows": states})

    def finish(self, action, status, details, result=None):
        receipt = {"id": self.id, "time": now().isoformat(), "task": action, "status": status, "duration": round(time.monotonic() - self.started, 2), "details": details, "result": result or {}, "evidence": str(self.evidence)}
        dump(self.evidence / "receipt.json", receipt)
        dump(log_path(self.settings, "History", self.id + ".json"), {k: v for k, v in receipt.items() if k != "result"})
        dump(self.folder / "receipt.json", receipt)
        return receipt


def history(settings):
    folder = log_path(settings, "History")
    if not folder.exists():
        return []
    rows = sorted((read(p) for p in folder.glob("*.json")), key=lambda r: r["time"], reverse=True)
    redirects = log_path(settings, "relocations.json")
    mappings = sorted(read(redirects), key=lambda item: len(item["from"]), reverse=True) if redirects.exists() else []
    for row in rows:
        original = row.get("evidence", "")
        for item in mappings:
            if original.casefold() == item["from"].casefold() or original.casefold().startswith(item["from"].casefold() + "\\"):
                row["evidence"] = item["to"] + original[len(item["from"]):]
                break
    return rows


def monitor_failure(folder, duration, diagnostic):
    folder = Path(folder)
    request = read(folder / "request.json")
    record = {"id": folder.name + "_monitor_" + uuid.uuid4().hex[:8], "time": now().isoformat(), "task": request["action"], "status": "failed", "duration": round(duration, 2), "details": "ตัวทำงานปิดก่อนตรวจสอบผล กรุณาตรวจบันทึกก่อนเริ่มซ้ำ", "result": {}, "evidence": str(folder)}
    dump(folder / "monitor_failure.json", {"receipt": record, "diagnostic": diagnostic})
    dump(log_path(request["settings"], "History", record["id"] + ".json"), {k: v for k, v in record.items() if k != "result"})
    return record
