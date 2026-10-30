"""Mailbox poll events and display state; no email actions are performed here."""
from datetime import datetime
from time import perf_counter
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import now


def scan_mail(outlook,settings,notify,task,workflow=True):
    common={"kind":"scan","task":task,"store":settings["store_name"],
            "folder":"/".join(settings["folder_path"])}
    notify({**common,"phase":"checking","time":now().isoformat(),"step":"กำลังตรวจเมล"})
    started=perf_counter()
    try:
        matches=outlook.scan_workflow(settings) if workflow else outlook.scan(settings)
    except Exception as error:
        detail=str(error) if isinstance(error,ValidationError) else "อ่านกล่องเมลไม่สำเร็จ"
        notify({**common,"phase":"error","time":now().isoformat(),"step":detail})
        if isinstance(error,ValidationError):
            raise
        raise ValidationError(detail) from error
    notify({**common,"phase":"checked","time":now().isoformat(),"matched":len(matches),
            "duration":round(perf_counter()-started,2),"step":f"พบ {len(matches)} เมลตรงเงื่อนไข"})
    return matches


class MonitorState:
    def __init__(self,workflows):
        self.rows={name:{"store":flow["store_name"],"folder":"/".join(flow["folder_path"]),
            "phase":"idle" if flow["enabled"] else "disabled","last_checked":None,"matched":None,
            "next_check":None,"step":"ยังไม่เริ่ม" if flow["enabled"] else "ปิดงาน"} for name,flow in workflows.items()}

    def apply(self,event):
        task=event.get("task")
        if task not in self.rows:
            return
        row=self.rows[task]
        if event["kind"]=="scan":
            row.update(phase=event["phase"],step=event["step"],next_check=None,
                       store=event["store"],folder=event["folder"])
            if event["phase"]=="checked":
                row.update(last_checked=event["time"],matched=event["matched"])
        elif event["kind"]=="poll_wait" and row["phase"]=="attention":
            row["next_check"]=event["next_check"]
        elif event["kind"]=="poll_wait" and row["phase"]!="error":
            row.update(phase="waiting",step="รอรอบถัดไป",next_check=event["next_check"])
        elif event["kind"]=="progress":
            row.update(phase="processing",step=event["step"],next_check=None)
        elif event["kind"]=="complete":
            row.update(phase="complete",step=event["receipt"]["details"])
        elif event["kind"]=="status":
            row["step"]=event["step"]
            if event.get("attention"):
                row.update(phase="attention",next_check=None)

    def stopped(self):
        for row in self.rows.values():
            if row["phase"]!="disabled":
                detail="หยุดตรวจเมล | "+row["step"] if row["phase"] in ("error","attention") else "หยุดตรวจเมล"
                row.update(phase="stopped",step=detail,next_check=None)

    def visible_step(self,task,at=None):
        row=self.rows[task]
        if row["phase"]=="attention" and row["next_check"]:
            remaining=(datetime.fromisoformat(row["next_check"])-(at or now())).total_seconds()
            return row["step"]+f" | ตรวจซ้ำใน {max(0,int(remaining))} วินาที"
        if row["phase"]=="waiting" and row["next_check"]:
            remaining=(datetime.fromisoformat(row["next_check"])-(at or now())).total_seconds()
            return f"รอรอบถัดไป {max(0,int(remaining))} วินาที" if remaining>0 else "รอคิวตรวจเมล"
        return row["step"]


def record_scan_issue(settings,event):
    """Retain an actionable error even when no job could be created."""
    from pathlib import Path
    from .vendor.fkp.persistence import dump
    path=Path(settings["logs"])/"Logs"/"last_scan_issue.json"
    dump(path,{"time":now().isoformat(),**event})
    return path
