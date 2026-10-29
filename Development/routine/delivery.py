"""Package only this job's verified files, preserving document subfolders."""
from pathlib import Path
import pyzipper

from .vendor.fkp.domain import ValidationError


def package_files(files, mail, job, name, relative_root=None):
    files = [Path(path) for path in files]
    if not files or any(not p.is_file() for p in files):
        raise ValidationError("ไฟล์ส่งงานไม่ครบ")
    if not mail.get("zip_attachment", True):
        if len(files) != 1:
            raise ValidationError("เอกสารหลายไฟล์ต้องเปิด ZIP เพื่อรักษาโครงสร้าง")
        return files[0]
    password = str(mail.get("zip_password", ""))
    if mail.get("zip_password_required", False) and not password:
        raise ValidationError("ระบุรหัส ZIP ก่อนส่งงาน")
    path = job.folder / (name + ".zip")
    entries = {}
    for file in files:
        entry = file.relative_to(relative_root).as_posix() if relative_root else file.name
        if entry.casefold() in {n.casefold() for n in entries}:
            raise ValidationError("ชื่อไฟล์แนบซ้ำ")
        entries[entry] = file
    with pyzipper.AESZipFile(path, "x", compression=pyzipper.ZIP_DEFLATED) as archive:
        if password:
            archive.setpassword(password.encode("utf-8"))
            archive.setencryption(pyzipper.WZ_AES, nbits=256)
        for entry, file in entries.items():
            archive.write(file, entry)
    with pyzipper.AESZipFile(path) as archive:
        if password:
            archive.setpassword(password.encode("utf-8"))
        if set(archive.namelist()) != set(entries):
            raise ValidationError("ไฟล์ใน ZIP ส่งงานไม่ครบ")
        for entry, file in entries.items():
            if archive.read(entry) != file.read_bytes():
                raise ValidationError("ตรวจ ZIP ส่งงานไม่ผ่าน")
    return path
