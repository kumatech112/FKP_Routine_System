"""SMB-compatible Windows leases held by an OS handle, not an orphan marker."""
from contextlib import contextmanager
from pathlib import Path
import json,os,re,socket,uuid
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import now,read,dump


def lock_path(folder,name):
    if not re.fullmatch(r'[A-Za-z0-9_-]+',name):
        raise ValueError('Invalid lock name')
    return Path(folder)/(name+'.lock')


def lock_notice(path):
    try:
        owner=read(path)
        who=str(owner.get('host') or 'ไม่ระบุเครื่อง (ล็อกรุ่นเก่า)')
        detail=f"เครื่อง {who}, PID {owner.get('pid','?')}, เริ่ม {owner.get('started','?')}"
    except Exception:
        detail='อ่านเจ้าของล็อกไม่ได้'
    return 'งานติดล็อก: '+detail+' | '+str(path)


@contextmanager
def lease(folder,name):
    import win32file,pywintypes
    path=lock_path(folder,name)
    path.parent.mkdir(parents=True,exist_ok=True)
    try:
        handle=win32file.CreateFile(str(path),win32file.GENERIC_READ|win32file.GENERIC_WRITE|0x10000,
            win32file.FILE_SHARE_READ,None,win32file.CREATE_NEW,
            win32file.FILE_ATTRIBUTE_NORMAL|0x04000000,None)  # DELETE_ON_CLOSE
    except (OSError,pywintypes.error) as error:
        if getattr(error,'winerror',None) in (32,33,80,183,303) or path.exists():
            raise ValidationError(lock_notice(path)) from error
        raise
    try:
        owner={'version':2,'host':socket.gethostname(),'pid':os.getpid(),
               'started':now().isoformat(),'token':uuid.uuid4().hex}
        win32file.WriteFile(handle,json.dumps(owner,ensure_ascii=False).encode('utf-8'))
        win32file.FlushFileBuffers(handle)
        yield
    finally:
        handle.Close()  # Windows/SMB also releases and deletes on process termination.


def recover_legacy_lock(folder,name,*,confirmed=False):
    """Explicit maintenance only: the user has stopped this task on every machine."""
    if not confirmed:
        raise ValidationError('ต้องยืนยันว่าหยุดงานประเภทนี้บนทุกเครื่องก่อนปลดล็อกเก่า')
    import win32file,pywintypes
    import ctypes
    from ctypes import wintypes
    path=lock_path(folder,name)
    try:
        handle=win32file.CreateFile(str(path),win32file.GENERIC_READ|win32file.GENERIC_WRITE|0x10000,
            win32file.FILE_SHARE_READ,None,win32file.OPEN_EXISTING,win32file.FILE_ATTRIBUTE_NORMAL,None)
    except (OSError,pywintypes.error) as error:
        if getattr(error,'winerror',None) in (2,3):
            raise ValidationError('ไม่พบล็อกเก่าของงานนี้') from error
        if getattr(error,'winerror',None) in (32,33,303):
            raise ValidationError('งานนี้ยังถือครองล็อกอยู่ จึงไม่ปลดล็อก') from error
        raise
    try:
        size=win32file.GetFileSize(handle)
        raw=win32file.ReadFile(handle,size)[1] if size else b''
        try:owner=json.loads(raw.decode('utf-8-sig'))
        except (ValueError,UnicodeError) as error:
            raise ValidationError('ข้อมูลล็อกเสียหาย จึงยังไม่ปลดล็อก: '+str(path)) from error
        if not isinstance(owner,dict) or not isinstance(owner.get('pid'),int) or not owner.get('started'):
            raise ValidationError('ไม่ใช่ล็อกรุ่นเก่าที่รองรับ จึงยังไม่ปลดล็อก')
        if owner.get('version') or owner.get('host'):
            raise ValidationError('ไม่ใช่ล็อกรุ่นเก่า จึงยังไม่ปลดล็อก')
        backup=path.parent/'RecoveredLocks'/(path.stem+'_'+now().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:8]+'.json')
        dump(backup,{'original_path':str(path),'owner':owner,'recovered_at':now().isoformat(),
                     'recovered_on':socket.gethostname(),'all_workers_stopped_confirmed':True})
        # Mark the exact open file for deletion; no close/unlink race with another worker.
        kernel=ctypes.WinDLL('kernel32',use_last_error=True)
        dispose=kernel.SetFileInformationByHandle
        dispose.argtypes=[wintypes.HANDLE,ctypes.c_int,ctypes.c_void_p,wintypes.DWORD]
        dispose.restype=wintypes.BOOL
        delete=ctypes.c_ubyte(1)
        if not dispose(int(handle),4,ctypes.byref(delete),ctypes.sizeof(delete)):
            raise ctypes.WinError(ctypes.get_last_error())
        return backup
    finally:handle.Close()
