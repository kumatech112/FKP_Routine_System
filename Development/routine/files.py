"""Exact file matching, private decryption and bounded ZIP/RAR extraction."""
from datetime import datetime
from pathlib import Path, PurePosixPath, PureWindowsPath
import os
import re
import shutil
import stat
import subprocess
import zipfile

from .rar_extract import extract_rar
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import copy_new, digest


def mail_date(match):
    token = match["mail_date"]
    return datetime.strptime(token, "%d/%m/%Y" if "/" in token else "%d%m%Y").date()


def find_archive(scan, match, body=""):
    folder = Path(scan["network_folder"])
    rule = re.compile(scan["network_file_regex"], re.IGNORECASE)
    month = (match["month_name"] + match["year"]).upper()
    found = []
    for item in folder.iterdir():
        result = rule.fullmatch(item.name)
        if item.is_file() and result and result.groupdict().get("month_token", "").upper() == month:
            found.append((item, result))
    # Transport date is exact for flexible intake; an old monthly mail cannot claim today's file.
    if match.get("received_date") or match.get("file_date"):
        stamp=match.get("file_date") or match["received_date"]
        explicit=[pair for pair in found if pair[0].name.casefold() in body.casefold()]
        found=explicit or [pair for pair in found if pair[1].groupdict().get("file_date")==stamp]
    # Prefer a filename explicitly present in Body, but never select an arbitrary one.
    from .triggers import logical_name
    requested=match.get("file_names",[])
    named = [(item, result) for item, result in found if item.name.casefold() in body.casefold()
             or any(logical_name(item.name)==logical_name(name) for name in requested)]
    if len(named)>1 and match.get("received_date"):
        dated=[pair for pair in named if pair[1].groupdict().get("file_date")==match["received_date"]]
        if dated:named=dated
    if named:
        found = named
    else:
        stamp = match.get("file_date") or match.get("received_date")
        if not stamp and match.get("mail_date"):
            stamp=mail_date(match).strftime("%y%m%d")
        exact = [(item, result) for item, result in found
                 if result.groupdict().get("file_date") == stamp]
        if exact or scan.get("require_exact_date_match", False):
            found = exact
    if not found:
        return None
    # Plain ZIP and encrypted ZIP with the same logical name represent the same delivery.
    groups = {}
    for item, result in found:
        base = re.sub(r"\.(gpg|pgp)$", "", item.name, flags=re.IGNORECASE).casefold()
        groups.setdefault(base, []).append(item)
    if len(groups) != 1:
        raise ValidationError("พบไฟล์ Performance หลายชุดตรงเมล ต้องตรวจสอบก่อนเลือก")
    candidates = next(iter(groups.values()))
    encrypted = [f for f in candidates if f.suffix.lower() in (".gpg", ".pgp")]
    if len(encrypted) > 1:
        raise ValidationError("พบทั้ง GPG และ PGP ต้องตรวจสอบไฟล์ก่อนเลือก")
    return (encrypted or candidates)[0]


def extract_payload(archive, job, decrypt, allow_pdf=False):
    staged = copy_new(archive, job.folder / Path(archive).name)
    if staged.suffix.lower() in (".gpg", ".pgp"):
        if not decrypt.get("enabled"):
            raise ValidationError("ยังไม่ได้เปิดถอด PGP")
        target = staged.with_suffix("")
        command = [decrypt["gpg_path"], "--batch", "--yes"]
        if decrypt.get("gpg_home"):
            command += ["--homedir", decrypt["gpg_home"]]
        passphrase = decrypt.get("passphrase", "")
        arguments = {}
        # Leave agent/default key-store handling intact when no override was saved.
        if passphrase:
            options = subprocess.run([decrypt["gpg_path"], "--dump-options"], capture_output=True,
                                     timeout=15, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            if options.returncode != 0 or b"--passphrase-fd" not in options.stdout.splitlines():
                raise ValidationError("ตัวถอด PGP ไม่รองรับรหัสผ่านที่บันทึกไว้")
            if b"--pinentry-mode" in options.stdout.splitlines():
                command += ["--pinentry-mode", "loopback"]
            command += ["--passphrase-fd", "0"]
            arguments["input"] = (passphrase + "\n").encode("utf-8")
        command += ["--output", str(target), "--decrypt", str(staged)]
        result = subprocess.run(command, capture_output=True, timeout=300,
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), **arguments)
        if result.returncode != 0 or not target.is_file():
            raise ValidationError("ถอด PGP ไม่สำเร็จ ตรวจสอบกุญแจของผู้ใช้และการตั้งค่าเดิม")
        staged = target
    job.check()
    folder = job.folder / "Extracted"
    if allow_pdf:
        with staged.open("rb") as stream:
            is_pdf = stream.read(1024).lstrip().startswith(b"%PDF-")
        if is_pdf:
            folder.mkdir(parents=True, exist_ok=True)
            name = staged.name if staged.suffix.casefold()==".pdf" else staged.stem+".pdf"
            copy_new(staged, folder / name)
            return folder
    if zipfile.is_zipfile(staged):
        extract_zip(staged, folder, password=decrypt.get("archive_password", ""))
    else:
        import rarfile
        if not rarfile.is_rarfile(staged):
            raise ValidationError("ไฟล์ที่ถอดได้ต้องเป็น ZIP หรือ RAR" + (" หรือ PDF" if allow_pdf else ""))
        extract_rar(staged, folder, check=job.check)
    return folder


def unpack_archive(archive, job, decrypt, excel_regex):
    folder = extract_payload(archive, job, decrypt)
    matches = [f for f in folder.rglob("*") if f.is_file()
               and re.fullmatch(excel_regex, f.name)]
    if len(matches) != 1:
        raise ValidationError("ต้องพบไฟล์ Daily ต้นทางที่ตรงเงื่อนไขเพียงไฟล์เดียว")
    return matches[0]


def extract_zip(source, folder, password=None):
    import pyzipper
    password=password.encode("utf-8") if isinstance(password,str) else password
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    with pyzipper.AESZipFile(source) as archive:
        members = archive.infolist()
        if len(members) > 10000 or sum(i.file_size for i in members) > 2 * 1024**3:
            raise ValidationError("ไฟล์ ZIP มีขนาดหรือจำนวนรายการเกินกำหนด")
        if any(member.flag_bits & 1 for member in members) and not password:
            raise ValidationError("ZIP ต้นทางมีรหัสผ่าน กรุณาระบุรหัส ZIP ในตั้งค่าอีเมล")
        seen = set()
        for member in members:
            name = member.filename.replace("\\", "/")
            parts = PurePosixPath(name)
            windows = PureWindowsPath(name)
            if (parts.is_absolute() or windows.drive or ".." in parts.parts
                    or ":" in name or stat.S_ISLNK(member.external_attr >> 16)):
                raise ValidationError("พบเส้นทางที่ไม่ปลอดภัยใน ZIP")
            identity = name.rstrip("/").casefold()
            if identity in seen:
                raise ValidationError("พบชื่อไฟล์ซ้ำใน ZIP")
            seen.add(identity)
            target = (folder / Path(*parts.parts)).resolve()
            if folder.resolve() not in target.parents:
                raise ValidationError("ไฟล์ ZIP ออกนอกโฟลเดอร์ Process")
        for member in members:
            target = folder / Path(*PurePosixPath(member.filename.replace("\\", "/")).parts)
            if member.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                try:
                    with archive.open(member,pwd=password) as incoming, target.open("xb") as outgoing:
                        shutil.copyfileobj(incoming,outgoing)
                except (RuntimeError,pyzipper.BadZipFile) as error:
                    raise ValidationError("แตก ZIP ต้นทางไม่สำเร็จ ตรวจรหัส ZIP หรือความสมบูรณ์ของไฟล์") from error


def make_attachment(stage02, mail, job):
    stage02 = Path(stage02)
    if not stage02.is_file() or not stage02.name.startswith("02_"):
        raise ValidationError("แนบเมลได้เฉพาะผลลัพธ์ขั้น 02")
    if not mail.get("zip_attachment", True):
        return stage02
    import pyzipper
    path = job.folder / (stage02.stem + ".zip")
    password = str(mail.get("zip_password", ""))
    with pyzipper.AESZipFile(path, "x", compression=pyzipper.ZIP_DEFLATED) as archive:
        if password:
            archive.setpassword(password.encode("utf-8"))
            archive.setencryption(pyzipper.WZ_AES, nbits=256)
        archive.write(stage02, arcname=stage02.name)
    with pyzipper.AESZipFile(path) as archive:
        if password:
            archive.setpassword(password.encode("utf-8"))
        if archive.namelist() != [stage02.name] or archive.read(stage02.name) != stage02.read_bytes():
            raise ValidationError("ตรวจ ZIP ขั้น 02 ไม่ผ่าน")
    return path


def archive_source(source, expected_hash, options):
    if not options.get("archive_processed_files", False):
        return None
    source = Path(source)
    if digest(source) != expected_hash:
        raise ValidationError("ไฟล์ต้นทางเปลี่ยนก่อนย้ายเข้า 03 Check")
    destination = Path(options["processed_archive_folder"]) / source.name
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.resolve() == source.resolve():
        raise ValidationError("ต้นทางและโฟลเดอร์เก็บไฟล์ซ้ำกัน")
    if destination.exists():
        if digest(destination) != expected_hash:
            raise ValidationError("มีไฟล์ชื่อเดิมแต่ข้อมูลต่างกันใน 03 Check")
    else:
        copy_new(source, destination)
    if digest(source) != expected_hash or digest(destination) != expected_hash:
        raise ValidationError("ตรวจสำเนาก่อนย้ายต้นทางไม่ผ่าน")
    source.unlink()
    return str(destination)
