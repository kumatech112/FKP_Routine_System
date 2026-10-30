"""Explain missing source files without replaying an already archived delivery."""
from pathlib import Path
from .triggers import logical_name,infer_sources


def missing_source_note(runtime,match,task,name=''):
    source=Path(runtime['scan']['network_folder'])
    archive=Path(runtime.get('archive',{}).get('processed_archive_folder') or source/'03 Check')
    if archive.is_dir():
        requested=match.get('file_names',[])
        if name and name!='รอชื่อไฟล์หรือไฟล์ที่ตรงเมล':requested=[name]
        if requested:
            wanted={logical_name(n) for n in requested}
            if any(p.is_file() and logical_name(p.name) in wanted for p in archive.iterdir()):
                return 'พบไฟล์ใน 03 Check แล้ว จึงไม่ส่งซ้ำจากโฟลเดอร์เก็บงาน'+(': '+name if name else '')
        elif task in ('customer_documents','enforcement'):
            try:
                if infer_sources(archive,match,task):return 'พบไฟล์ใน 03 Check แล้ว จึงไม่ส่งซ้ำจากโฟลเดอร์เก็บงาน'
            except Exception:pass
    return 'รอไฟล์ '+(name if name and name!='รอชื่อไฟล์หรือไฟล์ที่ตรงเมล' else 'ที่ตรงเมล')+' ใน '+str(source)
