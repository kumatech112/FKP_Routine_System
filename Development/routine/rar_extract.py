"""RAR metadata preflight and CRC-checked streams; never let UnRAR write archive paths."""
from pathlib import Path, PurePosixPath, PureWindowsPath
from threading import Event, Timer
import rarfile

from .config import resource_root
from .vendor.fkp.domain import ValidationError


def extract_rar(source, folder, check=lambda: None):
    folder = Path(folder)
    tool = resource_root() / "runtime" / "unrar" / "UnRAR.exe"
    if not tool.is_file():
        raise ValidationError("ไม่พบตัวแตก RAR ในชุดโปรแกรม")
    rarfile.UNRAR_TOOL = str(tool)
    try:
        rarfile.tool_setup(unar=False, bsdtar=False, sevenzip=False, sevenzip2=False)
        with rarfile.RarFile(source, errors="strict") as archive:
            members = archive.infolist()
            if archive.needs_password():
                raise ValidationError("RAR ชั้นในมีรหัสผ่าน กรุณาตรวจไฟล์ต้นทาง")
            if len(archive.volumelist()) != 1:
                raise ValidationError("ยังไม่รองรับ RAR แบบแบ่งหลายพาร์ต")
            if len(members) > 10000 or sum(m.file_size for m in members) > 2 * 1024**3:
                raise ValidationError("ไฟล์ RAR มีขนาดหรือจำนวนรายการเกินกำหนด")
            seen = set()
            targets = []
            for member in members:
                name = member.filename.replace("\\", "/")
                parts = PurePosixPath(name)
                windows = PureWindowsPath(name)
                if (parts.is_absolute() or windows.drive or ".." in parts.parts or ":" in name
                        or not parts.parts or member.is_symlink() or getattr(member,"file_redir",None)
                        or any(x.endswith((" ",".")) for x in parts.parts)):
                    raise ValidationError("พบเส้นทางที่ไม่ปลอดภัยใน RAR")
                identity = name.rstrip("/").casefold()
                target = (folder / Path(*parts.parts)).resolve()
                if identity in seen or folder.resolve() not in target.parents:
                    raise ValidationError("พบชื่อซ้ำหรือเส้นทางนอก Process ใน RAR")
                seen.add(identity)
                targets.append(target)
            folder.mkdir(parents=True, exist_ok=True)
            for member, target in zip(members, targets):
                check()
                if member.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as incoming, target.open("xb") as outgoing:
                    expired = Event()
                    # rarfile 4.2 owns this child; kill only that child if decompression stalls.
                    process = getattr(incoming, "_proc", None)
                    def timeout():
                        expired.set()
                        if process is not None and process.poll() is None:
                            process.kill()
                    timer = Timer(300, timeout)
                    timer.daemon = True
                    timer.start()
                    try:
                        count = 0
                        while True:
                            check()
                            block = incoming.read(1024 * 1024)
                            if not block:
                                break
                            count += len(block)
                            if count > member.file_size:
                                raise ValidationError("ขนาดข้อมูล RAR ไม่ตรงกับรายการ")
                            outgoing.write(block)
                        if expired.is_set() or count != member.file_size:
                            raise ValidationError("แตก RAR ไม่ครบหรือเกินเวลาที่กำหนด")
                    finally:
                        timer.cancel()
    except (rarfile.Error, OSError) as error:
        raise ValidationError("แตก RAR ไม่สำเร็จ ไฟล์เสียหายหรือรูปแบบไม่รองรับ") from error
