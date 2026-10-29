"""Workflow classification: Body-first documents, image-derived Enforcement, legacy Payment."""
from pathlib import Path
import re

from .files import mail_date
from .vendor.fkp.domain import ValidationError

DOCUMENT = re.compile(
    r"(?<![A-Za-z0-9])(?:\d{6}_)?A1015L?(?:_A1015L?)?_Document_Round"
    r"(?P<round>\d+)_(?P<doc_date>\d{8})\.zip\.(?:gpg|pgp)(?![A-Za-z0-9])", re.I)
ENFORCEMENT = re.compile(
    r"(?<![A-Za-z0-9])(?:\d{6}_)?A1015L?(?:_A1015L?)?_(?:Approval|Reject)_Enforcement"
    r"(?:_[A-Za-z][A-Za-z0-9-]*)*_(?P<doc_date>\d{8})(?:_\d+)?(?:\(\d+\))?"
    r"(?:\.(?:pdf|zip)(?:\.(?:gpg|pgp))?|\.(?:gpg|pgp))(?![A-Za-z0-9_.])", re.I)


def body_names(body, task):
    from .triggers import names
    return [item['name'] for item in names(body,task)]


def locate_body_file(folder, name):
    from .triggers import canonical_name,logical_name
    requested=canonical_name(name)
    exact=[f for f in Path(folder).iterdir() if f.is_file() and f.name.casefold()==requested.casefold()]
    if exact:return exact[0]
    stamp=re.match(r'^(\d{6})_',requested)
    found=[f for f in Path(folder).iterdir() if f.is_file() and logical_name(f.name)==logical_name(requested)
        and (not stamp or f.name.startswith(stamp[0]))]
    # Equivalent encrypted/plain copies are one payload; two transport dates are not.
    bases={re.sub(r'\.(gpg|pgp)$','',f.name,flags=re.I).casefold() for f in found}
    if len(bases)==1:
        encrypted=[f for f in found if f.suffix.casefold() in ('.gpg','.pgp')]
        found=encrypted or found
    if len(found)>1:raise ValidationError('ชื่อในเมลตรงกับหลายไฟล์ ต้องมีวันที่หรือชื่อไฟล์ที่ชัดเจน')
    return found[0] if found else None


def payment_file(scan, match):
    rule = re.compile(scan["network_file_regex"], re.I)
    found = []
    for path in Path(scan["network_folder"]).iterdir():
        result = rule.fullmatch(path.name)
        if path.is_file() and result and int(result["round"]) == int(match["round"]):
            if result["file_date"] == mail_date(match).strftime("%y%m%d"):
                found.append(path)
    bases = {}
    for path in found:
        base = re.sub(r"\.(gpg|pgp)$", "", path.name, flags=re.I).casefold()
        bases.setdefault(base, []).append(path)
    if len(bases) > 1:
        raise ValidationError("พบ Payment หลายไฟล์ตรงรอบและวันที่")
    if not bases:
        return None
    options = next(iter(bases.values()))
    encrypted = [p for p in options if p.suffix.lower() in (".gpg", ".pgp")]
    if len(encrypted) > 1:
        raise ValidationError("พบทั้ง GPG และ PGP ของรอบเดียวกัน")
    return (encrypted or options)[0]


def enforcement_match(subject, body):
    """Recognize decorated/near titles or an exact archive reference in Body."""
    from difflib import SequenceMatcher
    import unicodedata

    def normalized(text):
        return "".join(c for c in unicodedata.normalize("NFC", text).casefold()
                       if c.isalnum() or unicodedata.category(c).startswith("M"))

    filename = ENFORCEMENT.search(body)
    if filename:
        return {"mail_date": filename["doc_date"], "matched_by": "body_filename"}
    core = normalized("ขออนุมัติยึดอสังหาริมทรัพย์")
    title = normalized(subject[:1000])
    if ("rejectenforcement" in title or "ไม่อนุมัติยึดอสังหาริมทรัพย์" in title
            or ("reject" in title and "อสังหา" in title)):
        return {"matched_by": "reject_subject"}
    if core in title:
        return {"matched_by": "subject"}
    if core in normalized(body):
        return {"matched_by": "body_text"}
    # Restrict approximate matching to titles carrying the actual work's key words.
    if "ยึด" in title and "อสังหา" in title:
        for width in range(len(core)-2, len(core)+3):
            for start in range(max(0, len(title)-width+1)):
                score = SequenceMatcher(None, core, title[start:start+width], autojunk=False).ratio()
                if score >= .90:
                    return {"matched_by": "similar_subject", "similarity": round(score, 3)}
    return None

SUCCESS_BODY = re.compile(
    r"(?<![A-Za-z0-9])(?:\d{6}_)?A10(?:15L?|XX)(?:_A10(?:15L?|XX))?_Performance_SAM_OA_"
    r"(?P<month_name>[A-Za-z]{3})(?P<year>\d{2})\.(?:zip|rar)(?:\.(?:gpg|pgp))?(?![A-Za-z0-9_.])",re.I)
SUCCESS_FILE = re.compile(
    r"^(?:(?P<file_date>\d{6})_)?A1015L?(?:_A1015L?)?_Performance_SAM_OA_"
    r"(?P<month_name>[A-Za-z]{3})(?P<year>\d{2})\.(?:zip|rar)(?:\.(?:gpg|pgp))?$",re.I)


def success_rate_match(subject,body):
    if not re.search(r"Performance\s+success\s+rate",subject+"\n"+body,re.I):
        return None
    named=list(SUCCESS_BODY.finditer(body))
    title=re.search(r"Performance\s+success\s+rate.*?\b([A-Za-z]{3})['’\s]*(\d{2})\b.*?as\s+of\s+(\d{2}/\d{2}/\d{4})",
                    subject,re.I)
    result={}
    if title:
        result.update(month_name=title[1],year=title[2],mail_date=title[3])
    if named:
        periods={(m["month_name"].casefold(),m["year"]) for m in named}
        if len(periods)!=1:
            raise ValidationError("Body Success Rate อ้างหลายเดือน ต้องตรวจชื่อไฟล์")
        month,year=next(iter(periods))
        if title and (title[1].casefold(),title[2])!=(month,year):
            raise ValidationError("เดือนในหัวเมลและชื่อไฟล์ Success Rate ไม่ตรงกัน")
        result.update(month_name=named[0]["month_name"],year=year,body_file_name=named[0][0])
    return result if result.get("month_name") else None


def success_rate_file(scan,match,body):
    folder=Path(scan["network_folder"])
    month=(match["month_name"]+match["year"]).casefold()
    candidates=[]
    for path in folder.iterdir():
        found=SUCCESS_FILE.fullmatch(path.name)
        if path.is_file() and found and (found["month_name"]+found["year"]).casefold()==month:
            candidates.append((path,found))
    explicit=[pair for pair in candidates if pair[0].name.casefold() in body.casefold()]
    if explicit:
        candidates=explicit
    else:
        received=match.get("received_date")
        if received:
            candidates=[pair for pair in candidates if pair[1]["file_date"] in (None,received)]
        elif match.get("mail_date"):
            from datetime import timedelta
            stamp=mail_date(match)
            dates={stamp.strftime("%y%m%d"),(stamp+timedelta(days=1)).strftime("%y%m%d")}
            candidates=[pair for pair in candidates if pair[1]["file_date"] is None or pair[1]["file_date"] in dates]
    if not candidates:
        return None
    logical={}
    for path,found in candidates:
        identity=re.sub(r"\.(gpg|pgp)$","",path.name,flags=re.I).casefold()
        logical.setdefault(identity,[]).append(path)
    if len(logical)!=1:
        raise ValidationError("พบ Success Rate หลายไฟล์ตรง Body/วันที่ ต้องตรวจสอบก่อนเลือก")
    options=next(iter(logical.values()))
    encrypted=[p for p in options if p.suffix.casefold() in (".gpg",".pgp")]
    if len(encrypted)>1:
        raise ValidationError("พบทั้ง GPG และ PGP ของ Success Rate ชุดเดียวกัน")
    return (encrypted or options)[0]
