"""Native GUI for all workflows; profiles belong to the current Windows user."""
from copy import deepcopy
from pathlib import Path
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .ui import Application as BaseApplication
from .settings import TASKS, default_settings, dependency_checks, save, task_config, validate
from .workflows import run, finalize_enforcement, save_review_cell
from .theme import date_time, status_label, geometry
from .config import resource_root
from .outlook import list_sender_accounts
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import dump, history, read


class Application(BaseApplication):
    def __init__(self, values, path):
        self.sender_loading=False
        super().__init__(values, path)
        self.after(1000,self.monitor_clock)
        self.iconbitmap(str(resource_root() / "assets" / "app.ico"))
        for parent in self.winfo_children():
            for widget in parent.winfo_children():
                if isinstance(widget, ttk.Label) and str(widget.cget("textvariable")) == str(self.status):
                    widget.configure(wraplength=850)
                if isinstance(widget, ttk.Label) and widget.cget("text") == "Performance Daily":
                    widget.configure(text="งานประจำ")
        bar = self.start_button.master
        ttk.Button(bar, text="ตรวจความพร้อม", command=self.check_ready).pack(side="left", padx=(18, 0))
        self.lock_button=ttk.Button(bar,text="ปลดล็อกเก่า",command=self.recover_old_lock)
        self.lock_button.pack(side="left",padx=6)
        ttk.Button(bar, text="ส่งงาน OCR ที่ค้าง", command=self.open_review).pack(side="left", padx=6)

    def table(self,parent,columns,headings):
        if parent==getattr(self,"main_page",None):
            columns=("time","task","status","rows","detail")
            headings=("เวลา","งาน","สถานะ","จำนวนรายการ","รายละเอียด")
            self.main_tabs=ttk.Notebook(self.main_page)
            self.main_tabs.grid(row=1,column=0,sticky="nsew")
            self.status_page=ttk.Frame(self.main_tabs,padding=6)
            self.result_page=ttk.Frame(self.main_tabs,padding=6)
            self.main_tabs.add(self.status_page,text="สถานะตรวจเมล")
            self.main_tabs.add(self.result_page,text="ผลการทำงาน")
            self.result_page.columnconfigure(0,weight=1)
            self.result_page.rowconfigure(0,weight=1)
            self.build_monitor()
            parent=self.result_page
        return super().table(parent,columns,headings)

    def build_history(self):
        super().build_history()
        self.task_flags = {k: tk.BooleanVar(value=self.config_values["workflows"][k]["enabled"]) for k in TASKS}
        choices = ttk.Frame(self.main_page)
        choices.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.task_controls = []
        for key, label in TASKS.items():
            control = ttk.Checkbutton(choices, text="Success Rate" if key=="performance_success_rate" else label, variable=self.task_flags[key])
            control.pack(side="left", padx=(0, 14))
            self.task_controls.append(control)
        self.main_page.rowconfigure(0, weight=0)
        self.main_page.rowconfigure(1, weight=1)
        self.main_table.grid(row=0, column=0, sticky="nsew")

    def build_monitor(self):
        from .monitoring import MonitorState
        self.monitor_state=MonitorState(self.config_values["workflows"])
        self.status_page.columnconfigure(0,weight=1)
        self.status_page.rowconfigure(1,weight=1)
        frame=self.table(self.status_page,("task","state","last","matched","store","folder"),
            ("งาน","สถานะ","ตรวจเมลล่าสุด","พบเมล","กล่องเมล","โฟลเดอร์"))
        frame.grid(row=0,column=0,sticky="ew",pady=(0,8))
        self.monitor_tree=frame.tree
        self.monitor_tree.configure(height=len(TASKS))
        for name,width in (("task",170),("state",250),("last",165),("matched",70),("store",245),("folder",170)):
            self.monitor_tree.column(name,width=width,stretch=name=="state",anchor="e" if name=="matched" else "w")
        log=ttk.LabelFrame(self.status_page,text="บันทึกการตรวจเมล",padding=8)
        log.grid(row=1,column=0,sticky="nsew")
        log.columnconfigure(0,weight=1);log.rowconfigure(0,weight=1)
        self.live_log=tk.Text(log,wrap="word",font=("Tahoma",11),state="disabled")
        scroll=ttk.Scrollbar(log,orient="vertical",command=self.live_log.yview)
        self.live_log.configure(yscrollcommand=scroll.set)
        self.live_log.grid(row=0,column=0,sticky="nsew")
        scroll.grid(row=0,column=1,sticky="ns")
        self.refresh_monitor()

    def refresh_monitor(self):
        if not hasattr(self,"monitor_tree"):
            return
        for task,row in self.monitor_state.rows.items():
            values=("Success Rate" if task=="performance_success_rate" else TASKS[task],self.monitor_state.visible_step(task),
                date_time(row["last_checked"]) if row["last_checked"] else "",
                "" if row["matched"] is None else row["matched"],row["store"],row["folder"])
            if self.monitor_tree.exists(task):
                self.monitor_tree.item(task,values=values)
            else:
                self.monitor_tree.insert("","end",iid=task,values=values)

    def monitor_clock(self):
        self.refresh_monitor()
        self.after(1000,self.monitor_clock)

    def append_live_log(self,event,message=None):
        from .vendor.fkp.persistence import now
        message=message or event.get("step","")
        if not message:
            return
        label=TASKS.get(event.get("task"),"ระบบ")
        stamp=date_time(event.get("time") or now().isoformat())
        self.live_log.configure(state="normal")
        self.live_log.insert("end",f"{stamp} [{label}] {message}\n")
        lines=int(self.live_log.index("end-1c").split(".")[0])
        if lines>500:
            self.live_log.delete("1.0",f"{lines-500+1}.0")
        self.live_log.see("end")
        self.live_log.configure(state="disabled")

    def load_sender_accounts(self):
        if self.sender_loading or (self.worker and self.worker.is_alive()):
            return
        self.sender_loading=True
        self.sender_refresh.state(["disabled"])
        self.status.set("กำลังอ่านบัญชีผู้ส่งจาก Outlook")
        def work():
            try:
                accounts=list_sender_accounts()
                self.events.put({"kind":"sender_accounts","accounts":accounts})
            except Exception:
                self.events.put({"kind":"sender_accounts","error":"อ่านบัญชี Outlook ไม่สำเร็จ"})
        threading.Thread(target=work,name="outlook-sender-accounts",daemon=True).start()

    def build_settings(self):
        self.settings_page.columnconfigure(0, weight=1)
        self.settings_page.rowconfigure(0, weight=1)
        inner = ttk.Notebook(self.settings_page)
        inner.grid(row=0, column=0, sticky="nsew")
        self.setting_controls, self.fields, self.flow_vars = [], {}, {}
        def page(title):
            outer = ttk.Frame(inner)
            inner.add(outer, text=title)
            outer.columnconfigure(0, weight=1)
            outer.rowconfigure(0, weight=1)
            canvas = tk.Canvas(outer, highlightthickness=0)
            scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=scroll.set)
            canvas.grid(row=0, column=0, sticky="nsew")
            scroll.grid(row=0, column=1, sticky="ns")
            body = ttk.Frame(canvas, padding=14)
            body.columnconfigure(1, weight=1)
            item = canvas.create_window((0, 0), window=body, anchor="nw")
            body.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.bind("<Configure>", lambda e: canvas.itemconfigure(item, width=e.width))
            return body
        def field(parent, row, label, var, browse=None, secret=False):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=5)
            entry = ttk.Entry(parent, textvariable=var, show="*" if secret else "")
            entry.grid(row=row, column=1, sticky="ew", pady=5)
            self.setting_controls.append(entry)
            if browse:
                button = ttk.Button(parent, text="เลือก", command=lambda: self.browse_value(var, browse))
                button.grid(row=row, column=2, padx=(6, 0), pady=5)
                self.setting_controls.append(button)
            return entry
        general = page("ทั่วไป")
        labels = {"logs_root": ("Log และ Process", "dir"), "output_root": ("งานสำเร็จ", "dir"),
            "payment_database": ("Access Payment", "file"), "database": ("Access Daily", "file"),
            "gpg_home": ("กุญแจ PGP (ว่าง = ค่าเครื่อง)", "dir"), "gpg_path": ("ตัวถอด PGP (GnuPG)", "file"),
            "ocr_path": ("โปรแกรม OCR", "file"), "poll": ("ตรวจเมลทุก (วินาที)", None),
            "scan_hours": ("ตรวจเมลย้อนหลัง (ชั่วโมง)",None), "scan_limit": ("จำนวนเมลสูงสุดต่อโฟลเดอร์",None)}
        for row, (key, (label, browse)) in enumerate(labels.items()):
            var = self.fields[key] = tk.StringVar(value=str(self.config_values[key]))
            field(general, row, label, var, browse)
        ttk.Label(general, text="โปรไฟล์ผู้ใช้").grid(row=len(labels), column=0, sticky="w")
        profile = ttk.Entry(general)
        profile.insert(0, str(self.settings_path))
        profile.configure(state="readonly")
        profile.grid(row=len(labels), column=1, columnspan=2, sticky="ew")
        email = page("อีเมล")
        self.mail_vars = {"send_from": tk.StringVar(value=self.config_values["mail"]["send_from"]),
            "body_prefix": tk.StringVar(value=self.config_values["mail"].get("body_prefix", "")),
            "mode": tk.StringVar(value="send" if self.config_values["mail"].get("send_immediately", True) else "draft"),
            "zip": tk.BooleanVar(value=self.config_values["mail"].get("zip_attachment", True))}
        modes = ttk.Frame(email)
        modes.grid(row=0, column=0, columnspan=3, sticky="w")
        for label, value in (("ส่งทันที", "send"), ("เปิด Draft", "draft")):
            control = ttk.Radiobutton(modes, text=label, variable=self.mail_vars["mode"], value=value)
            control.pack(side="left", padx=(0, 12))
            self.setting_controls.append(control)
        control = ttk.Checkbutton(modes, text="แนบเป็น ZIP", variable=self.mail_vars["zip"])
        control.pack(side="left")
        self.setting_controls.append(control)
        sender=ttk.Frame(email)
        sender.grid(row=1,column=0,columnspan=3,sticky="ew",pady=5)
        sender.columnconfigure(0,weight=1)
        ttk.Label(sender,text="บัญชีผู้ส่ง (Send From)").grid(row=0,column=0,sticky="w")
        controls=ttk.Frame(sender)
        controls.grid(row=1,column=0,sticky="ew",pady=(3,0))
        controls.columnconfigure(0,weight=1)
        self.sender_combo=ttk.Combobox(controls,textvariable=self.mail_vars["send_from"])
        self.sender_combo.grid(row=0,column=0,sticky="ew")
        self.sender_refresh=ttk.Button(controls,text="เลือกบัญชี",command=self.load_sender_accounts)
        self.sender_refresh.grid(row=0,column=1,padx=(6,0))
        self.setting_controls.extend((self.sender_combo,self.sender_refresh))
        field(email, 2, "ข้อความนำ", self.mail_vars["body_prefix"])
        self.secret_vars = {k: tk.StringVar(value=self.config_values["secrets"].get(k, "")) for k in ("passphrase", "zip_password")}
        secret_entries = [field(email, 3, "รหัสกุญแจ PGP (ไม่บังคับ)", self.secret_vars["passphrase"], secret=True),
            field(email, 4, "รหัส ZIP ส่งงาน", self.secret_vars["zip_password"], secret=True)]
        show = tk.BooleanVar(value=False)
        ttk.Checkbutton(email, text="แสดงรหัส", variable=show,
            command=lambda: [e.configure(show="" if show.get() else "*") for e in secret_entries]).grid(row=5, column=1, sticky="w")
        self.recipient_vars = {}
        for number, key in enumerate(("payment_online", "performance", "customer_documents", "enforcement", "performance_success_rate")):
            group = self.config_values["mail"].get("workflow_recipients", {}).get(key, {"to": [], "cc": []})
            self.recipient_vars[key] = {f: tk.StringVar(value="; ".join(group.get(f, []))) for f in ("to", "cc")}
            field(email, 6+2*number, TASKS[key]+" To", self.recipient_vars[key]["to"])
            field(email, 7+2*number, TASKS[key]+" CC", self.recipient_vars[key]["cc"])
        rules = page("รับงาน")
        for number, (key, label) in enumerate(TASKS.items()):
            flow = self.config_values["workflows"][key]
            group = ttk.LabelFrame(rules, text=label, padding=12)
            group.grid(row=number, column=0, columnspan=3, sticky="ew", pady=(0, 10))
            group.columnconfigure(1, weight=1)
            values = self.flow_vars[key] = {"store_name": tk.StringVar(value=flow["store_name"]),
                "folder_path": tk.StringVar(value="/".join(flow["folder_path"])),
                "network_folder": tk.StringVar(value=flow["network_folder"]),
                "extra_inbox": tk.BooleanVar(value=flow.get("extra_inbox", False))}
            field(group, 0, "กล่องเมล", values["store_name"])
            field(group, 1, "โฟลเดอร์ Outlook", values["folder_path"])
            field(group, 2, "โฟลเดอร์ไฟล์ต้นทาง", values["network_folder"], "dir")
            check = ttk.Checkbutton(group, text="ตรวจ Inbox เพิ่มเติม", variable=values["extra_inbox"])
            check.grid(row=3, column=1, sticky="w")
            self.setting_controls.append(check)
        actions = ttk.Frame(self.settings_page)
        actions.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.save_button = ttk.Button(actions, text="บันทึก", command=self.save)
        self.save_button.pack(side="left")
        self.reset_button = ttk.Button(actions, text="คืนค่าเส้นทาง", command=self.reset)
        self.reset_button.pack(side="left", padx=6)

    def browse_value(self, variable, kind):
        current = variable.get()
        result = filedialog.askdirectory(initialdir=current) if kind=="dir" else filedialog.askopenfilename(initialdir=str(Path(current).parent))
        if result:
            variable.set(result)

    def candidate(self):
        values = deepcopy(self.config_values)
        for key, var in self.fields.items():
            values[key] = int(var.get()) if key in ("poll","scan_hours","scan_limit") else var.get()
        values["secrets"] = {k: v.get() for k, v in self.secret_vars.items()}
        values["mail"].update(send_from=self.mail_vars["send_from"].get(), body_prefix=self.mail_vars["body_prefix"].get(),
            send_immediately=self.mail_vars["mode"].get()=="send", zip_attachment=self.mail_vars["zip"].get())
        for key, group in self.recipient_vars.items():
            values["mail"].setdefault("workflow_recipients", {})[key] = {
                f: [s.strip() for s in var.get().split(";") if s.strip()] for f, var in group.items()}
        for key, group in self.flow_vars.items():
            for name, var in group.items():
                value = var.get()
                values["workflows"][key][name] = [s.strip() for s in value.split("/") if s.strip()] if name=="folder_path" else value
            values["workflows"][key]["enabled"] = self.task_flags[key].get()
        return values

    def dirty(self):
        try:
            return self.candidate()!=self.config_values
        except (ValueError, tk.TclError):
            return True

    def save(self):
        try:
            candidate = validate(self.candidate())
            save(candidate, self.settings_path)
            self.config_values = candidate
            self.refresh_history()
            self.status.set("บันทึกโปรไฟล์ผู้ใช้แล้ว")
        except Exception as error:
            self.status.set(str(error) if isinstance(error, (ValidationError, ValueError)) else "บันทึกโปรไฟล์ไม่สำเร็จ")

    def reset(self):
        defaults = default_settings()
        for key, var in self.fields.items():
            var.set(str(defaults[key]))
        self.status.set("คืนค่าเส้นทางเริ่มต้นแล้ว ยังไม่ได้บันทึก")

    def busy(self, value):
        super().busy(value)
        if hasattr(self,"lock_button"):
            self.lock_button.configure(state="disabled" if value else "normal")
        for control in self.task_controls:
            control.configure(state="disabled" if value else "normal")

    def start(self):
        if self.sender_loading:
            self.status.set("รออ่านบัญชีผู้ส่งให้เสร็จก่อนเริ่ม")
            return
        if self.worker and self.worker.is_alive():
            return
        if self.dirty():
            self.status.set("บันทึกการตั้งค่าก่อนเริ่มงาน")
            return
        self.stop_event.clear()
        self.busy(True)
        self.status.set("กำลังตรวจงานที่เลือก")
        values = deepcopy(self.config_values)
        from .monitoring import MonitorState
        self.monitor_state=MonitorState(values["workflows"])
        self.refresh_monitor()
        self.append_live_log({},"เริ่มเฝ้าตรวจเมล")
        def work():
            try:
                run(values, self.stop_event, self.events.put)
            except BaseException as error:
                self.events.put({"kind": "status", "step": str(error) if isinstance(error, ValidationError) else "เริ่มงานไม่สำเร็จ: "+type(error).__name__})
            finally:
                self.events.put({"kind": "stopped"})
        self.worker = threading.Thread(target=work, daemon=False)
        self.worker.start()

    def recover_old_lock(self):
        if self.worker and self.worker.is_alive():
            self.status.set("หยุดงานก่อนปลดล็อกเก่า")
            return
        selected=self.monitor_tree.selection()
        task=selected[0] if selected and selected[0] in TASKS else 'customer_documents'
        folder=Path(task_config(self.config_values,task)['logs'])/'Logs/Locks'
        name='performance' if task=='performance' else 'workflow'
        prompt=(TASKS[task]+'\n'+str(folder/(name+'.lock'))+
                '\n\nหยุดงานประเภทนี้บนทุกเครื่องแล้วใช่ไหม?\nระบบจะสำรองและปลดเฉพาะล็อกรุ่นเก่า โดยไม่ส่งเมลหรือเปลี่ยนประวัติงาน')
        if not messagebox.askyesno('ปลดล็อกเก่า',prompt):
            return
        self.busy(True)
        def work():
            try:
                from .locking import recover_legacy_lock
                backup=recover_legacy_lock(folder,name,confirmed=True)
                message='สำรองและปลดล็อกเก่าแล้ว กดเริ่มเพื่อตรวจเมล | '+str(backup)
            except Exception as error:
                message=str(error) if isinstance(error,ValidationError) else 'ปลดล็อกเก่าไม่สำเร็จ ('+type(error).__name__+')'
            self.events.put({'kind':'status','task':task,'step':message})
            self.events.put({'kind':'stopped'})
        self.worker=threading.Thread(target=work,daemon=False)
        self.worker.start()

    def check_ready(self):
        if self.worker and self.worker.is_alive():
            self.status.set("หยุดงานก่อนตรวจความพร้อม")
            return
        values = deepcopy(self.config_values)
        def work():
            results = []
            for key, flow in values["workflows"].items():
                if flow["enabled"]:
                    problems = dependency_checks(values, key, write_probe=True)
                    results.append(TASKS[key]+": "+("; ".join(problems) if problems else "ตรวจส่วนประกอบผ่าน"))
            self.events.put({"kind": "status", "step": "\n".join(results)})
        threading.Thread(target=work, daemon=True).start()

    def refresh_history(self):
        self.history_rows = []
        unreadable = []
        for key in TASKS:
            try:
                self.history_rows.extend(history(task_config(self.config_values, key)))
            except Exception:
                unreadable.append(TASKS[key])
        self.render_history()
        if unreadable:
            self.status.set("อ่านประวัติไม่ได้: " + ", ".join(unreadable))

    def poll_events(self):
        try:
            while True:
                self.handle_event(self.events.get_nowait())
        except queue.Empty:
            pass
        if self.closing and (not self.worker or not self.worker.is_alive()):
            self.destroy()
            return
        self.after(150,self.poll_events)

    def handle_event(self,event):
        self.monitor_state.apply(event)
        kind=event["kind"]
        if kind=="sender_accounts":
            self.sender_loading=False
            self.sender_refresh.state(["!disabled"])
            if event.get("error"):
                self.status.set(event["error"])
            else:
                accounts=event["accounts"]
                self.sender_combo.configure(values=accounts)
                self.status.set(f"พบบัญชีผู้ส่งใน Outlook {len(accounts)} บัญชี")
        elif kind=="scan":
            self.status.set(TASKS.get(event["task"],event["task"])+": "+event["step"])
            self.progress.grid_remove()
            self.append_live_log(event)
        elif kind=="poll_wait":
            self.append_live_log(event,"รอตรวจเมลรอบถัดไป "+date_time(event["next_check"]))
        elif kind in ("progress","status"):
            self.status.set(event["step"])
            if kind=="progress":
                self.percent.set(event["percent"])
                self.progress.grid()
            else:
                self.progress.grid_remove()
            self.append_live_log(event)
        elif kind=="complete":
            row=event["receipt"]
            self.results.append(row)
            self.main_table.tree.insert("",0,iid=row["id"],values=(date_time(row["time"]),row["task"],
                status_label(row["status"]),row.get("result",{}).get("records",row.get("result",{}).get("cells","")),row["details"]))
            self.status.set(row["details"])
            self.progress.grid_remove()
            self.append_live_log(event,row["details"])
            self.refresh_history()
        elif kind=="review_done":
            dialog=event.get("dialog")
            if dialog and dialog.winfo_exists():
                if event.get("error"):dialog.message.set(event["error"])
                else:dialog.destroy()
        elif kind=="stopped":
            self.busy(False)
            self.progress.grid_remove()
            self.monitor_state.stopped()
            self.append_live_log(event,"หยุดเฝ้าตรวจเมล")
            if self.stop_event.is_set():
                self.status.set("หยุดการทำงานแล้ว")
        self.refresh_monitor()

    def open_review(self):
        if self.worker and self.worker.is_alive():
            self.status.set("หยุดตรวจเมลก่อนส่งงานที่ค้าง")
            return
        selected=self.main_table.tree.selection()
        receipt=next((r for r in self.results if selected and r["id"]==selected[0]),None)
        if self.notebook.index(self.notebook.select())==1:
            selected=self.history_table.tree.selection()
            row=next((r for r in self.filtered if selected and r["id"]==selected[0]),None)
            if row and (Path(row["evidence"])/"receipt.json").is_file():
                receipt=read(Path(row["evidence"])/"receipt.json")
        path=Path(receipt.get("result",{}).get("review","")) if receipt else None
        if not path or not path.is_file():
            self.status.set("เลือกรายการ Enforcement ที่ OCR แล้วและยังไม่ส่ง")
            return
        context=read(path.parent/"review_context.json")
        if context.get("workflow_version",0)<4:
            self.status.set("รายการนี้ไม่รองรับการส่งงาน OCR ที่ค้าง")
            return
        self.stop_event.clear()
        self.busy(True)
        values=deepcopy(self.config_values)
        def notify(event):
            self.events.put({**event,"task":"enforcement"})
        def work():
            try:
                finalize_enforcement(path,{},notify,self.stop_event,values=values)
            except Exception as error:
                notify({"kind":"status","step":str(error) if isinstance(error,ValidationError)
                    else "ส่งงาน OCR ที่ค้างไม่สำเร็จ กรุณาตรวจบันทึก"})
            finally:
                self.events.put({"kind":"stopped"})
        self.worker=threading.Thread(target=work,daemon=False)
        self.worker.start()

    def close(self):
        if self.worker and self.worker.is_alive():
            if messagebox.askyesno("FKP Routine System", "หยุดงานและปิดเมื่อถึงจุดที่หยุดได้หรือไม่?"):
                self.stop_event.set()
                self.closing = True
            return
        if self.dirty() and not messagebox.askyesno("FKP Routine System", "ปิดโดยไม่บันทึกการตั้งค่าที่แก้ไขหรือไม่?"):
            return
        self.config_values["geometry"] = self.geometry()
        try:
            save(self.config_values, self.settings_path)
        except Exception:
            self.status.set("บันทึกโปรไฟล์ก่อนปิดไม่สำเร็จ")
            return
        self.destroy()


class ReviewDialog(tk.Toplevel):
    def __init__(self, owner, path):
        super().__init__(owner)
        self.owner, self.path, self.review = owner, path, read(path)
        self.title("Enforcement Review")
        geometry(self, "1100x720+20+20")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.message = tk.StringVar()
        ttk.Label(self, textvariable=self.message).grid(row=0, column=0, sticky="w", padx=12, pady=8)
        frame = owner.table(self, ("page","row","field","value","score","state"), ("หน้า","แถว","ฟิลด์","ข้อความ","OCR","สถานะ"))
        frame.grid(row=1, column=0, sticky="nsew", padx=12)
        self.tree = frame.tree
        self.tree.bind("<Double-1>", lambda _: self.edit())
        bar = ttk.Frame(self)
        bar.grid(row=2, column=0, sticky="ew", padx=12, pady=10)
        ttk.Button(bar, text="เปิดภาพต้นฉบับ", command=self.image).pack(side="left")
        ttk.Button(bar, text="ตรวจ/แก้ช่องที่เลือก", command=self.edit).pack(side="left", padx=6)
        ttk.Button(bar, text="ยืนยันและส่ง ZIP / Draft", command=self.finalize).pack(side="right")
        self.render()

    def render(self):
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for cell in self.review["cells"]:
            self.tree.insert("", "end", iid=cell["id"], values=(cell["page_id"],cell["row"],cell["field"],
                cell["value"][:120],cell["confidence"],"ตรวจแล้ว" if cell["state"]=="reviewed" else "รอตรวจ"))
        pending = sum(c["state"]!="reviewed" for c in self.review["cells"])
        self.message.set(f"รอตรวจ {pending} จาก {len(self.review['cells'])} ช่อง/หน้า รวมช่องว่างและหัวตาราง")
        if selected:
            self.tree.selection_set(selected)

    def selected(self):
        selection = self.tree.selection()
        return next((c for c in self.review["cells"] if selection and c["id"]==selection[0]), None)

    def image(self):
        cell = self.selected()
        if cell and Path(cell["image"]).is_file():
            os.startfile(cell["image"])

    def edit(self):
        cell = self.selected()
        if cell is None:
            return
        dialog = tk.Toplevel(self)
        dialog.title(cell["field"])
        dialog.geometry("720x420")
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=1)
        text = tk.Text(dialog, wrap="word", font=("Tahoma",11))
        text.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        text.insert("1.0", cell["value"])
        def commit():
            try:
                self.review = save_review_cell(self.path, cell["id"], text.get("1.0", "end-1c"))
            except ValidationError as error:
                self.message.set(str(error))
                return
            dialog.destroy()
            self.render()
        ttk.Button(dialog, text="ยืนยันว่าตรวจช่องนี้แล้ว", command=commit).grid(row=1, column=0, pady=(0,12))

    def finalize(self):
        if any(c["state"]!="reviewed" for c in self.review["cells"]):
            self.message.set("ยังตรวจไม่ครบทุกช่องและหัวตาราง")
            return
        self.owner.stop_event.clear()
        self.owner.busy(True)
        def work():
            error = None
            try:
                finalize_enforcement(self.path, {}, self.owner.events.put, self.owner.stop_event, values=self.owner.config_values)
            except Exception as failure:
                error = str(failure) if isinstance(failure, ValidationError) else "เก็บงานไม่สำเร็จ กรุณาตรวจบันทึก"
            finally:
                self.owner.events.put({"kind": "review_done", "dialog": self, "error": error})
                self.owner.events.put({"kind": "stopped"})
        self.owner.worker = threading.Thread(target=work, daemon=False)
        self.owner.worker.start()
