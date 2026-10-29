"""Private, windowless Tesseract calls; retain its native Thai text spacing."""
from functools import lru_cache
from pathlib import Path
from statistics import mean
import csv
import io
import os
import subprocess
import tempfile
import unicodedata
from .vendor.fkp.domain import ValidationError

def hidden_process_options():
    options={'creationflags':getattr(subprocess,'CREATE_NO_WINDOW',0)}
    if os.name=='nt':
        startup=subprocess.STARTUPINFO()
        startup.dwFlags|=subprocess.STARTF_USESHOWWINDOW
        startup.wShowWindow=subprocess.SW_HIDE
        options['startupinfo']=startup
    return options

def environment(executable):
    env=dict(os.environ)
    data=Path(executable).parent/'tessdata'
    if data.is_dir():env['TESSDATA_PREFIX']=str(data)
    return env

@lru_cache(maxsize=4)
def _check_languages(executable,stamp):
    result=subprocess.run([executable,'--list-langs'],capture_output=True,timeout=15,
        env=environment(executable),**hidden_process_options())
    languages=set(result.stdout.decode('utf-8',errors='replace').splitlines())
    if result.returncode or not {'tha','eng'}<=languages:
        raise ValidationError('OCR ต้องมีภาษาไทยและอังกฤษ')

def configure_ocr(executable):
    executable=Path(executable)
    if not executable.is_file():raise ValidationError('ไม่พบ Tesseract OCR')
    _check_languages(str(executable),executable.stat().st_mtime_ns)
    return executable

def recognize(image,executable,psm=6,lang='tha+eng',whitelist=''):
    executable=configure_ocr(executable)
    # TXT preserves Thai word/combining-mark boundaries; TSV is used only for confidence.
    with tempfile.TemporaryDirectory(prefix='FKP_OCR_') as private:
        source=Path(private)/'cell.png';base=Path(private)/'result'
        image.save(source)
        command=[str(executable),str(source),str(base),'-l',lang,'--oem','1','--psm',str(psm)]
        if whitelist:command+=['-c','tessedit_char_whitelist='+whitelist]
        command+=['-c','tessedit_create_txt=1','-c','tessedit_create_tsv=1']
        result=subprocess.run(command,capture_output=True,timeout=90,
            env=environment(executable),**hidden_process_options())
        if result.returncode or not base.with_suffix('.txt').is_file():
            raise ValidationError('อ่าน OCR ไม่สำเร็จ กรุณาตรวจไฟล์และตัว OCR')
        text=base.with_suffix('.txt').read_text(encoding='utf-8').strip()
        text=unicodedata.normalize('NFC',text).replace('\u0e4d\u0e32','\u0e33')
        scores=[float(row['conf']) for row in csv.DictReader(io.StringIO(
            base.with_suffix('.tsv').read_text(encoding='utf-8')),delimiter='\t')
            if row.get('text','').strip() and float(row['conf'])>=0]
    return text,round(mean(scores),1) if scores else 0
