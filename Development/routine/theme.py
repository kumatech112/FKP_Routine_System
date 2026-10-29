"""Shared native Windows sizing, fonts and date formatting."""
from datetime import datetime
import re
import tkinter.font as font
from tkinter import ttk

BODY_FONT_PT = ("Tahoma", 11)
TITLE_FONT_PT = ("Segoe UI", 18, "bold")
STATUS_FONT_PT = ("Segoe UI", 10, "bold")


def configure(window):
    font.nametofont("TkDefaultFont").configure(family=BODY_FONT_PT[0], size=BODY_FONT_PT[1])
    font.nametofont("TkTextFont").configure(family=BODY_FONT_PT[0], size=BODY_FONT_PT[1])
    style = ttk.Style(window)
    style.configure("Treeview", font=BODY_FONT_PT, rowheight=32)
    style.configure("Treeview.Heading", font=(*BODY_FONT_PT, "bold"))
    style.configure("Title.TLabel", font=TITLE_FONT_PT)
    style.configure("Status.TLabel", font=STATUS_FONT_PT)


def geometry(window, saved):
    window.update_idletasks()
    width, height = window.winfo_screenwidth(), window.winfo_screenheight()
    left, top, right, bottom = 0, 0, width, height
    try:
        import ctypes
        from ctypes import wintypes
        area = wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(area), 0):
            left, top, right, bottom = area.left, area.top, area.right, area.bottom
    except Exception:
        pass
    match = re.fullmatch(r"(\d+)x(\d+)([+-]\d+)?([+-]\d+)?", saved or "")
    w, h, x, y = (int(match[1]), int(match[2]), int(match[3] or 20), int(match[4] or 20)) if match else (1180, 780, 20, 20)
    w, h = min(max(700, w), right-left), min(max(480, h), bottom-top)
    x, y = max(left, min(x, right-w)), max(top, min(y, bottom-h))
    window.geometry(f"{w}x{h}{x:+d}{y:+d}")
    window.minsize(min(1000, right-left), min(650, bottom-top))


def date_time(value):
    try:
        return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M:%S")
    except (ValueError, TypeError):
        return str(value)


def status_label(value):
    return {"success": "สำเร็จ", "failed": "ไม่สำเร็จ", "cancelled": "ยกเลิก", "draft_opened": "เปิด Draft แล้ว", "needs_review": "รอตรวจทาน"}.get(value, value)
