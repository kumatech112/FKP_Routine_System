"""Owned STA Office sessions and saved-cache filtering; never refresh a connection."""
from contextlib import contextmanager
import gc
import re

from .domain import ValidationError, key


@contextmanager
def apartment():
    import pythoncom
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        gc.collect()
        pythoncom.CoUninitialize()


class Excel:
    def __enter__(self):
        import win32com.client
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.AutomationSecurity = 3
        self.app.EnableEvents = False
        self.app.AskToUpdateLinks = False
        self.books = []
        try:
            self.books.append(self.app.Workbooks.Add())
            self.app.Calculation = -4135
        except BaseException:
            self.app.Quit()
            raise
        return self

    def open(self, path, readonly=False):
        book = self.app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=readonly, IgnoreReadOnlyRecommended=True)
        self.books.append(book)
        for connection in book.Connections:
            try:
                connection.OLEDBConnection.EnableRefresh = False
                connection.OLEDBConnection.RefreshOnFileOpen = False
            except Exception:
                pass
        return book

    def calculate(self):
        self.app.CalculateFull()

    def close(self, book):
        book.Close(SaveChanges=False)
        self.books.remove(book)

    def __exit__(self, kind, error, traceback):
        failures = []
        for book in reversed(self.books):
            try:
                book.Close(SaveChanges=False)
            except Exception as failure:
                failures.append(failure)
        self.books.clear()
        try:
            self.app.Quit()
        finally:
            self.app = None
        if failures and error is None:
            raise ValidationError("ปิดไฟล์ทำงานไม่สำเร็จ กรุณาตรวจสอบ Excel") from failures[0]


def field_key(value):
    return re.sub(r"[^A-Z0-9]", "", key(value))


class CacheFilters:
    def __init__(self, book, sheet):
        self.pivots = list(book.Worksheets(sheet).PivotTables())
        self.pages = []
        self.applied = {}
        for pivot in self.pivots:
            for field in pivot.PageFields:
                name = field_key(field.Name)
                self.pages.append((pivot, field, name))

    def apply(self, desired, job):
        normalized = {field_key(k): v for k, v in desired.items()}
        seen = set()
        for pivot in self.pivots:
            job.check()
            changed = False
            pivot.ManualUpdate = True
            try:
                for owner, field, name in self.pages:
                    if owner.Name != pivot.Name or name not in normalized:
                        continue
                    seen.add(name)
                    value = normalized[name]
                    if name == "OACODE" and re.fullmatch(r"\d{4}-.+", str(value)):
                        value = str(value).split("-", 1)[0]
                    identity = (pivot.Name, name)
                    signature = tuple(key(v) for v in value) if isinstance(value, list) else key(value).replace("(ALL)", "ALL")
                    if identity not in self.applied:
                        self.applied[identity] = key(str(field.CurrentPage)).replace("(ALL)", "ALL")
                    if self.applied[identity] == signature:
                        continue
                    changed = True
                    field.ClearAllFilters()
                    field.EnableMultiplePageItems = False
                    if isinstance(value, list):
                        items = list(field.PivotItems())
                        wanted = {key(v) for v in value}
                        if not wanted.issubset({key(str(i.Name)) for i in items}):
                            raise ValidationError("ไม่พบกลุ่มผลิตภัณฑ์ที่เลือก")
                        field.EnableMultiplePageItems = True
                        for item in items:
                            if key(str(item.Name)) in wanted:
                                item.Visible = True
                        for item in items:
                            if key(str(item.Name)) not in wanted:
                                item.Visible = False
                    elif key(value) not in ("ALL", "(ALL)"):
                        try:
                            field.CurrentPage = value
                        except Exception as error:
                            raise ValidationError("ไม่พบตัวกรอง " + str(value)) from error
                    if not isinstance(value, list) and key(value) not in ("ALL", "(ALL)") and key(str(field.CurrentPage)) != key(value):
                        raise ValidationError("ตรวจสอบตัวกรองต้นทางไม่ผ่าน")
                    self.applied[identity] = signature
            finally:
                pivot.ManualUpdate = False
            if changed:
                pivot.Update()
        missing = set(normalized) - seen
        if missing:
            raise ValidationError("ต้นทางไม่มีตัวกรอง " + ", ".join(sorted(missing)))


def matrix(sheet, address):
    value = sheet.Range(address).Value2
    if not isinstance(value, tuple):
        return ((value,),)
    return value
