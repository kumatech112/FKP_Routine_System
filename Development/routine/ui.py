"""Native three-page GUI. Starting is an explicit production action."""
from pathlib import Path
import csv
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .config import default_config, load_config, load_legacy
from .service import monitor
from .theme import configure, date_time, geometry, status_label
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import dump, history


class Application(tk.Tk):
    def __init__(self, config, path):
        super().__init__()
        self.title("FKP Routine System")
        self.config_values, self.settings_path = dict(config), Path(path)
        self.events, self.stop_event = queue.Queue(), threading.Event()
        self.worker = None
        self.closing = False
        self.status = tk.StringVar(value="พร้อมทำงาน")
        self.percent = tk.IntVar(value=0)
        self.fields = {k: tk.StringVar(value=str(v)) for k, v in config.items() if k != "geometry"}
        self.results, self.history_rows, self.filtered = [], [], []
        configure(self)
        geometry(self, config["geometry"])
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        ttk.Label(self, text="FKP Routine System", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=16, pady=(10, 4))
        toolbar = ttk.Frame(self)
        toolbar.grid(row=1, column=0, sticky="ew", padx=16)
        self.start_button = ttk.Button(toolbar, text="เริ่มทำงาน", command=self.start)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(toolbar, text="หยุด", command=self.stop_event.set, state="disabled")
        self.stop_button.pack(side="left", padx=(6, 0))
        ttk.Label(toolbar, text="Performance Daily").pack(side="right")
        strip = ttk.Frame(self)
        strip.grid(row=2, column=0, sticky="ew", padx=16, pady=6)
        strip.columnconfigure(0, weight=1)
        ttk.Label(strip, textvariable=self.status, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(strip, variable=self.percent, maximum=100, length=180)
        self.progress.grid(row=0, column=1, padx=8)
        self.progress.grid_remove()
        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=3, column=0, sticky="nsew", padx=12, pady=(0, 8))
        self.main_page, self.history_page, self.settings_page = [ttk.Frame(self.notebook, padding=10) for _ in range(3)]
        for page, label in zip((self.main_page, self.history_page, self.settings_page),
                               ("หน้าหลัก", "ประวัติการทำงาน", "ตั้งค่า")):
            self.notebook.add(page, text=label)
        self.main_page.columnconfigure(0, weight=1)
        self.main_page.rowconfigure(0, weight=1)
        self.main_table = self.table(self.main_page, ("time", "status", "rows", "detail"),
                                    ("เวลา", "สถานะ", "จำนวนแถว", "รายละเอียด"))
        self.main_table.grid(row=0, column=0, sticky="nsew")
        self.build_history()
        self.build_settings()
        self.refresh_history()
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(150, self.poll_events)

    def table(self, parent, columns, headings):
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for name, title in zip(columns, headings):
            tree.heading(name, text=title)
            tree.column(name, width=180 if name in ("time", "task") else 130, minwidth=45,
                        anchor="e" if name in ("rows", "duration") else "w",
                        stretch=name == "detail")
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.tree = tree
        return frame

    def build_history(self):
        self.history_page.columnconfigure(0, weight=1)
        self.history_page.rowconfigure(1, weight=1)
        bar = ttk.Frame(self.history_page)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(bar, text="ค้นหา").pack(side="left")
        self.search = tk.StringVar()
        ttk.Entry(bar, textvariable=self.search, width=28).pack(side="left", padx=6)
        self.search.trace_add("write", lambda *_: self.render_history())
        ttk.Button(bar, text="รีเฟรช", command=self.refresh_history).pack(side="left")
        ttk.Button(bar, text="ส่งออกผลที่แสดง", command=self.export_history).pack(side="right")
        self.history_table = self.table(self.history_page, ("time", "task", "status", "duration", "detail"),
                                       ("เวลา", "งาน", "สถานะ", "ระยะเวลา (วินาที)", "รายละเอียด"))
        self.history_table.grid(row=1, column=0, sticky="nsew")
        self.history_table.tree.bind("<Double-1>", self.open_evidence)
        self.sort_column, self.sort_reverse = "time", True
        for col in ("time", "task", "status", "duration", "detail"):
            self.history_table.tree.heading(col, command=lambda name=col: self.sort_history(name))

    def build_settings(self):
        self.settings_page.columnconfigure(0, weight=1)
        self.settings_page.rowconfigure(0, weight=1)
        canvas = tk.Canvas(self.settings_page, highlightthickness=0)
        scroll = ttk.Scrollbar(self.settings_page, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        body = ttk.Frame(canvas, padding=14)
        item = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(item, width=e.width))
        body.columnconfigure(0, weight=1)
        group = ttk.LabelFrame(body, text="ไฟล์และโฟลเดอร์", padding=12)
        group.grid(row=0, column=0, sticky="ew")
        group.columnconfigure(1, weight=1)
        labels = {"legacy_config": "การตั้งค่า Routine เดิม", "logs": "Log และ Process",
                  "output": "งานสำเร็จ", "database": "ฐาน Daily (Access)"}
        self.setting_controls = []
        for row, (key, label) in enumerate(labels.items()):
            ttk.Label(group, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
            entry = ttk.Entry(group, textvariable=self.fields[key])
            entry.grid(row=row, column=1, sticky="ew", pady=4)
            button = ttk.Button(group, text="เลือก", command=lambda name=key: self.browse(name))
            button.grid(row=row, column=2, padx=(6, 0), pady=4)
            self.setting_controls.extend([entry, button])
        email = ttk.LabelFrame(body, text="อีเมล", padding=12)
        email.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.mail_summary = tk.StringVar(value="")
        ttk.Label(email, textvariable=self.mail_summary, wraplength=800).pack(anchor="w")
        actions = ttk.Frame(self.settings_page)
        actions.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.save_button = ttk.Button(actions, text="บันทึก", command=self.save)
        self.save_button.pack(side="left")
        self.reset_button = ttk.Button(actions, text="คืนค่าเริ่มต้น", command=self.reset)
        self.reset_button.pack(side="left", padx=6)
        self.update_mail_summary()

    def update_mail_summary(self):
        try:
            mail = load_legacy(self.config_values)["mail"]
            mode = "ส่งทันที" if mail.get("send_immediately", True) else "เปิด Draft"
            self.mail_summary.set(f"ผู้ส่ง: {mail['send_from']}\nTo: {'; '.join(mail['to'])}   CC: {'; '.join(mail['cc'])}\n{mode}   Performance Daily แนบขั้น 01" + (" (ZIP)" if mail.get("zip_attachment") else ""))
        except Exception:
            self.mail_summary.set("ยังอ่านการตั้งค่าเมลเดิมไม่ได้")

    def browse(self, key):
        current = self.fields[key].get()
        selected = filedialog.askopenfilename(initialdir=str(Path(current).parent)) if key in ("legacy_config", "database") else filedialog.askdirectory(initialdir=current)
        if selected:
            self.fields[key].set(selected)

    def dirty(self):
        return any(v.get() != str(self.config_values[k]) for k, v in self.fields.items())

    def save(self):
        candidate = {**self.config_values, **{k: v.get() for k, v in self.fields.items()}}
        try:
            # Validate before replacing the actual settings file.
            import tempfile
            with tempfile.TemporaryDirectory() as folder:
                temporary = Path(folder) / "settings.json"
                dump(temporary, candidate)
                candidate = load_config(temporary)
                load_legacy(candidate)
            dump(self.settings_path, candidate)
            self.config_values = candidate
            self.update_mail_summary()
            self.refresh_history()
            self.status.set("บันทึกการตั้งค่าแล้ว")
        except Exception as error:
            self.status.set(str(error) if isinstance(error, ValidationError) else "บันทึกการตั้งค่าไม่สำเร็จ")

    def reset(self):
        for k, value in default_config().items():
            if k in self.fields:
                self.fields[k].set(value)
        self.status.set("คืนค่าเริ่มต้นแล้ว ยังไม่ได้บันทึก")

    def busy(self, value):
        self.start_button.configure(state="disabled" if value else "normal")
        self.stop_button.configure(state="normal" if value else "disabled")
        for control in self.setting_controls + [self.save_button, self.reset_button]:
            control.configure(state="disabled" if value else "normal")

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        if self.dirty():
            self.status.set("บันทึกการตั้งค่าก่อนเริ่มงาน")
            return
        self.stop_event.clear()
        self.busy(True)
        self.status.set("กำลังตรวจเมล Performance")
        config = dict(self.config_values)
        def work():
            try:
                monitor(config, self.stop_event, self.events.put)
            except BaseException as error:
                self.events.put({"kind": "status", "step": str(error) if isinstance(error, ValidationError) else "เริ่มงานไม่สำเร็จ: " + type(error).__name__})
            finally:
                self.events.put({"kind": "stopped"})
        self.worker = threading.Thread(target=work, daemon=False)
        self.worker.start()

    def poll_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event["kind"] in ("progress", "status"):
                    self.status.set(event["step"])
                    if event["kind"] == "progress":
                        self.percent.set(event["percent"])
                        self.progress.grid()
                    else:
                        self.progress.grid_remove()
                elif event["kind"] == "complete":
                    row = event["receipt"]
                    self.results.append(row)
                    self.main_table.tree.insert("", 0, values=(date_time(row["time"]), status_label(row["status"]),
                        row.get("result", {}).get("records", ""), row["details"]))
                    self.status.set(row["details"])
                    self.progress.grid_remove()
                    self.refresh_history()
                elif event["kind"] == "stopped":
                    self.busy(False)
                    self.progress.grid_remove()
                    if self.stop_event.is_set():
                        self.status.set("หยุดการทำงานแล้ว")
        except queue.Empty:
            pass
        if self.closing and (not self.worker or not self.worker.is_alive()):
            self.destroy()
            return
        self.after(150, self.poll_events)

    def refresh_history(self):
        try:
            self.history_rows = history({"logs": self.config_values["logs"]})
        except Exception:
            self.history_rows = []
        self.render_history()

    def sort_history(self, column):
        self.sort_reverse = not self.sort_reverse if self.sort_column == column else False
        self.sort_column = column
        self.render_history()

    def render_history(self):
        query = self.search.get().casefold()
        self.filtered = [r for r in self.history_rows if query in " ".join(
            [date_time(r["time"]), r["task"], r["status"], r["details"]]).casefold()]
        field = "details" if self.sort_column == "detail" else self.sort_column
        self.filtered.sort(key=lambda row: row.get(field, ""), reverse=self.sort_reverse)
        tree = self.history_table.tree
        selected = set(tree.selection())
        tree.delete(*tree.get_children())
        for row in self.filtered:
            tree.insert("", "end", iid=row["id"], values=(
                date_time(row["time"]), row["task"], status_label(row["status"]), row["duration"], row["details"]))
        remaining = selected & set(tree.get_children())
        if remaining:
            tree.selection_set(tuple(remaining))

    def export_history(self):
        if not self.filtered:
            self.status.set("ไม่มีประวัติในขอบเขตที่แสดง")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            with Path(path).open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["เวลา", "งาน", "สถานะ", "ระยะเวลา", "รายละเอียด"])
                for row in self.filtered:
                    writer.writerow([date_time(row["time"]), row["task"], row["status"], row["duration"], row["details"]])
            self.status.set(f"ส่งออกประวัติ {len(self.filtered)} รายการแล้ว")
        except OSError:
            self.status.set("ส่งออกประวัติไม่สำเร็จ")

    def open_evidence(self, event=None):
        selected = self.history_table.tree.selection()
        if selected:
            row = next((r for r in self.filtered if r["id"] == selected[0]), None)
            if row and Path(row["evidence"]).is_dir():
                os.startfile(row["evidence"])

    def close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("FKP Routine System", "หยุดงานและปิดโปรแกรมเมื่อถึงจุดที่หยุดได้หรือไม่?"):
                return
            self.closing = True
            self.stop_event.set()
            self.status.set("กำลังหยุดงาน")
            return
        if self.dirty() and not messagebox.askyesno("FKP Routine System", "ปิดโดยไม่บันทึกการตั้งค่าที่แก้ไขหรือไม่?"):
            return
        self.config_values["geometry"] = self.geometry()
        dump(self.settings_path, self.config_values)
        self.destroy()
