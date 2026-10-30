"""Outlook adapter; no mail is sent merely by loading settings or scanning."""
from dataclasses import dataclass
from datetime import datetime, timedelta
from html import escape
import re

from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import BANGKOK


@dataclass(frozen=True)
class MailRef:
    entry_id: str
    store_id: str
    subject: str
    body: str
    match: dict


class Outlook:
    def __init__(self):
        import win32com.client
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.session = self.app.GetNamespace("MAPI")

    def scan(self, settings):
        store = next((s for s in self.session.Stores
                      if str(s.DisplayName).casefold() == settings["store_name"].casefold()), None)
        if store is None:
            account = next((a for a in self.session.Accounts if str(a.SmtpAddress).casefold() == settings["store_name"].casefold()), None)
            if account is not None:
                store = account.DeliveryStore
        if store is None:
            raise ValidationError("ไม่พบกล่อง Outlook ตามที่ตั้งไว้")
        folder = store.GetRootFolder()
        for number, name in enumerate(settings["folder_path"]):
            try:
                folder = store.GetDefaultFolder(6) if number == 0 and name.casefold() in ("inbox", "กล่องจดหมายเข้า") else folder.Folders.Item(name)
            except Exception as error:
                raise ValidationError("ไม่พบหรือไม่มีสิทธิ์โฟลเดอร์ Outlook: " + name) from error
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        subject_rule = re.compile(settings["subject_regex"], re.IGNORECASE)
        earliest = datetime.now(BANGKOK) - timedelta(hours=float(settings.get("received_within_hours", 24)))
        result = []
        for index in range(1, min(items.Count, int(settings.get("max_scan_items", 200))) + 1):
            mail = items.Item(index)
            if int(getattr(mail, "Class", 0)) != 43:
                continue
            received = mail.ReceivedTime
            if received.tzinfo is None:
                received = received.replace(tzinfo=BANGKOK)
            if received < earliest:
                break
            if settings.get("unread_only", True) and not mail.UnRead:
                continue
            subject, body = str(mail.Subject), str(mail.Body)
            if settings.get("flexible_trigger"):
                from .triggers import match_mail
                collection=getattr(mail,"Attachments",None)
                attachments=[str(collection.Item(i).FileName) for i in range(1,getattr(collection,"Count",0)+1)]
                groups=match_mail(settings['flexible_trigger'],subject,body,attachments,
                    received.astimezone(BANGKOK).strftime('%y%m%d'),settings['subject_regex'])
            elif settings.get("processor") == "enforcement":
                from .matching import enforcement_match
                groups = enforcement_match(subject, body)
            elif settings.get("processor") == "performance_success_rate":
                from .matching import success_rate_match
                groups = success_rate_match(subject,body)
                if groups is not None:
                    groups["received_date"]=received.astimezone(BANGKOK).strftime("%y%m%d")
            else:
                match = subject_rule.fullmatch(subject)
                groups = match.groupdict() if match else None
            if groups is not None:
                result.append(MailRef(str(mail.EntryID), str(store.StoreID), subject, body, groups))
        return list(reversed(result))

    def original(self, reference):
        return self.session.GetItemFromID(reference.entry_id, reference.store_id)

    def prepare_forward(self, reference, attachment, settings):
        if not settings.get("to") or not str(settings.get("send_from", "")).strip():
            raise ValidationError("ระบุผู้รับและบัญชีผู้ส่งก่อนส่งงาน")
        if not attachment.is_file():
            raise ValidationError("ไม่พบไฟล์แนบ")
        sender = str(settings["send_from"]).strip()
        matches = [a for a in self.session.Accounts
                   if str(a.SmtpAddress or "").strip().casefold()==sender.casefold()]
        if len(matches)!=1:
            raise ValidationError("ไม่พบบัญชีผู้ส่งที่เลือกเพียงบัญชีเดียวใน Outlook")
        original = self.original(reference)
        forwarded = original.Forward()
        try:
            while forwarded.Recipients.Count:
                forwarded.Recipients.Remove(1)
            for kind, names in ((1, settings["to"]), (2, settings["cc"])):
                for name in names:
                    forwarded.Recipients.Add(name).Type = kind
            if not forwarded.Recipients.ResolveAll():
                raise ValidationError("ตรวจสอบชื่อผู้รับหรือกลุ่ม Outlook ไม่ผ่าน")
            require_sender(forwarded,self.session.Accounts,sender)
            while forwarded.Attachments.Count:
                forwarded.Attachments.Remove(1)
            forwarded.Attachments.Add(str(attachment.resolve()))
            if forwarded.Attachments.Count != 1:
                raise ValidationError("ไฟล์แนบเมลไม่ครบ")
            prefix = str(settings.get("body_prefix", ""))
            if prefix:
                forwarded.HTMLBody = "<p>" + escape(prefix).replace("\n", "<br>") + "</p>" + forwarded.HTMLBody
            return forwarded
        except BaseException:
            forwarded.Close(1)
            raise

    def deliver(self, forwarded, settings):
        require_sender(forwarded,self.session.Accounts,settings["send_from"])
        if settings.get("send_immediately", True):
            forwarded.Send()
            return "sent"
        forwarded.Display(False)
        try:
            forwarded.Save()
            verify_sender(forwarded,settings["send_from"])
        except BaseException:
            forwarded.Close(1)
            raise
        return "draft_opened"

    def complete(self, reference, scan):
        original = self.original(reference)
        if scan.get("mark_as_read", False):
            original.UnRead = False
            original.Save()
        if scan.get("open_email_window", False):
            original.Display(False)


    def scan_workflow(self, settings):
        from copy import deepcopy
        found = {r.entry_id: r for r in self.scan(settings)}
        if settings.get("extra_inbox") and len(settings["folder_path"]) > 1:
            inbox = deepcopy(settings)
            inbox["folder_path"] = [settings["folder_path"][0]]
            found.update({r.entry_id: r for r in self.scan(inbox)})
        return list(found.values())

    def inline_images(self, reference, folder):
        from pathlib import Path
        import re
        original = self.original(reference)
        html = str(original.HTMLBody or "")
        folder = Path(folder)
        folder.mkdir(parents=True, exist_ok=True)
        files = []
        for index in range(1, original.Attachments.Count+1):
            item = original.Attachments.Item(index)
            suffix = Path(str(item.FileName)).suffix.casefold()
            if suffix not in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"):
                continue
            try:
                cid = str(item.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F"))
            except Exception:
                cid = ""
            if cid and "cid:"+cid not in html:
                continue
            path = folder / f"image_{index}{suffix}"
            if not path.exists():
                item.SaveAsFile(str(path))
            files.append(path)
        return files

def list_sender_accounts():
    """Read current-profile SMTP addresses only after the UI asks."""
    import win32com.client
    from .vendor.fkp.office import apartment
    with apartment():
        outlook=win32com.client.Dispatch("Outlook.Application")
        addresses={str(a.SmtpAddress or "").strip() for a in outlook.Session.Accounts}
        return sorted((a for a in addresses if a),key=str.casefold)


def require_sender(mail,accounts,sender):
    sender=str(sender or "").strip()
    if not sender:
        raise ValidationError("เลือกบัญชีผู้ส่งในตั้งค่าอีเมล")
    matches=[a for a in accounts if str(a.SmtpAddress or "").strip().casefold()==sender.casefold()]
    if len(matches)!=1:
        raise ValidationError("ไม่พบบัญชีผู้ส่งที่เลือกเพียงบัญชีเดียวใน Outlook")
    mail.SendUsingAccount=matches[0]
    verify_sender(mail,sender)


def verify_sender(mail,sender):
    selected=getattr(mail.SendUsingAccount,"SmtpAddress",None)
    if str(selected or "").strip().casefold()!=str(sender).strip().casefold():
        raise ValidationError("Outlook ไม่ยืนยันบัญชีผู้ส่งที่เลือก จึงหยุดก่อนส่ง")
