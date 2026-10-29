"""Shared mail classification. Fuzzy words identify work; office/date/round stay exact."""
from datetime import datetime
from difflib import SequenceMatcher
import html
import re
import unicodedata

from .vendor.fkp.domain import ValidationError

MONTHS={name.lower():name for name in ('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')}
PREFIX=r'(?<![A-Za-z0-9])(?:(?P<file_date>\d{6})_)?(?P<office>A10(?:15|xx)L?)(?:_A10(?:15|xx)L?)?_'
ENDING=r'(?:\.(?:zip|rar|pdf|xlsm|xlsx))?(?:\.(?:gpg|pgp))?(?![A-Za-z0-9_.])'
RULES={
 'customer_documents':re.compile(PREFIX+r'Document_Round(?P<round>\d+)_(?P<doc_date>\d{8})'+ENDING,re.I),
 'payment_online':re.compile(PREFIX+r'TSS_PAYMENT_Online(?P<round>\d+)'+ENDING,re.I),
 'performance':re.compile(PREFIX+r'Daily_Recovery_Performance_(?P<month_name>[A-Za-z]{3})(?P<year>\d{2})'+ENDING,re.I),
 'performance_success_rate':re.compile(PREFIX+r'Performance_SAM_OA_(?P<month_name>[A-Za-z]{3})(?P<year>\d{2})'+ENDING,re.I),
 'enforcement':re.compile(PREFIX+r'(?P<decision>Approval|Reject)_Enforcement(?P<bank>(?:_[A-Za-z][A-Za-z0-9-]*)*)_(?P<doc_date>\d{8})(?P<variant>(?:_\d+)?(?:\(\d+\))?)'+ENDING,re.I),
}


def clean(text):
    text=html.unescape(unicodedata.normalize('NFKC',str(text or '')))
    text=re.sub('[\u200b-\u200f\ufeff]','',text)
    return re.sub(r'\s*_\s*','_',text)


def canonical_name(name):
    return re.sub(r'A10xx', 'A1015',clean(name),flags=re.I)


def valid_date(token):
    try:
        stamp=datetime.strptime(token,'%d/%m/%Y' if '/' in token else '%d%m%Y')
        return stamp.strftime('%d%m%Y') if 2000<=stamp.year<=2100 else None
    except ValueError:return None


def dates(text):
    return list(dict.fromkeys(v for t in re.findall(r'(?<!\d)(?:\d{2}/\d{2}/20\d{2}|\d{4}20\d{2})(?!\d)',text) if (v:=valid_date(t))))


def names(text,task):
    result={}
    for found in RULES[task].finditer(clean(text)):
        data={k:v for k,v in found.groupdict().items() if v is not None}
        if data.get('doc_date') and not valid_date(data['doc_date']):continue
        if data.get('month_name') and data['month_name'].lower() not in MONTHS:continue
        if task=='payment_online' and data.get('round') and not 1<=int(data['round'])<=12:continue
        data['name']=canonical_name(found[0]);data['office']=canonical_name(data['office']).upper()
        if 'doc_date' in data:data['mail_date']=data['doc_date']
        if 'month_name' in data:data['month_name']=MONTHS[data['month_name'].lower()]
        result.setdefault(data['name'].casefold(),data)
    return list(result.values())


def related(text,task):
    text=clean(text).casefold()
    tokens=set(re.findall(r'[a-z]+',text))
    def word(w):return w in tokens or any(abs(len(w)-len(t))<=2 and SequenceMatcher(None,w,t).ratio()>=.84 for t in tokens)
    if task=='payment_online':return word('payment') and word('online')
    if task=='performance':return word('daily') and word('recovery') and word('performance') and not (word('adjust') or word('success'))
    if task=='performance_success_rate':return word('performance') and word('success') and word('rate')
    if task=='enforcement':
        from .matching import enforcement_match
        return bool(enforcement_match(text,'')) or (word('enforcement') and (word('approval') or word('reject')))
    if word('document') and (word('round') or 'เอกสาร' in text):return True
    compact=''.join(c for c in text if c.isalnum() or unicodedata.category(c).startswith('M'))
    key='เอกสารลูกค้า'
    if key in compact:return True
    if not any(k in compact for k in ('เอกสาร','ลูกค้า','ลูกคา')):return False
    return any(SequenceMatcher(None,key,compact[i:i+n]).ratio()>=.84
        for n in range(len(key)-1,len(key)+2) for i in range(max(0,len(compact[:500])-n+1)))


def one(values,label):
    values=list(dict.fromkeys(values))
    if len(values)>1:raise ValidationError('เมลอ้าง '+label+' หลายค่า จึงยังเลือกไฟล์ไม่ได้')
    return values[0] if values else None


def match_mail(task,subject,body,attachments=(),received_date=None,legacy_rule=None):
    subject,body=clean(subject),clean(body)
    if re.match(r'(?i)^\s*(?:(?:FW|FWD|RE):\s*)*(?:Recall:|เรียกคืน)',subject):return None
    sources=[('Body',body),('ชื่อไฟล์แนบ','\n'.join(attachments)),('Subject',subject)]
    selected=[];matched_by=None
    body_tasks={other for other in RULES if names(body,other)}
    if body_tasks and task not in body_tasks:return None
    for label,text in sources:
        selected=names(text,task)
        if selected:matched_by=label+' / ชื่อไฟล์';break
    # An explicit different office cannot be rescued by a generic title.
    office_text=body+'\n'+'\n'.join(attachments)
    if not names(office_text,task) and re.search(r'A10(?!15)\d{2}L?_',office_text,re.I):return None
    # Adjust is never a Daily source, even if its title was copied from Daily.
    if task=='performance' and re.search(r'\badjust(?:ment)?\b',subject+'\n'+body,re.I):return None
    if not selected:
        # Strong names for another workflow outrank a copied/mistyped subject.
        if any(names(body,other) for other in RULES if other!=task):return None
        for label,text in [('Subject',subject),('Body',body[:8000])]:
            if related(text,task):matched_by=label+' / คำสำคัญหรือข้อความคล้าย';break
        if not matched_by:
            legacy=re.fullmatch(legacy_rule,subject,re.I) if legacy_rule and legacy_rule!='.*' else None
            return legacy.groupdict() if legacy else None
    result={'matched_by':matched_by,'file_names':[x['name'] for x in selected]}
    if received_date:result['received_date']=received_date
    try:
        if selected:
            for key,label in [('round','รอบ'),('month_name','เดือน'),('year','ปี')]:
                vals=[x[key] for x in selected if key in x]
                if vals and task not in ('customer_documents','enforcement'):result[key]=one(vals,label)
            if len(selected)==1:result.update({k:v for k,v in selected[0].items() if k!='name'})
            # Filename identity is authoritative; a wrong subject date is not substituted.
            file_dates=[x['mail_date'] for x in selected if x.get('mail_date')]
            if len(set(file_dates))==1:result['mail_date']=file_dates[0]
        if task=='payment_online' and selected and result.get('file_date'):
            result['mail_date']=datetime.strptime(result['file_date'],'%y%m%d').strftime('%d%m%Y')
        if 'mail_date' not in result:
            result['mail_date']=one(dates(body),'วันที่ใน Body') or one(dates(subject),'วันที่ใน Subject')
        if 'round' not in result and task in ('customer_documents','payment_online') and not selected:
            result['round']=one(re.findall(r'(?:Round|รอบ)\s*[_:-]?\s*(\d+)(?!\d)',body,re.I),'รอบใน Body') or one(re.findall(r'(?:Round|รอบ)\s*[_:-]?\s*(\d+)(?!\d)',subject,re.I),'รอบใน Subject')
        if task in ('performance','performance_success_rate') and not result.get('month_name'):
            periods=[]
            for text in (body,subject):
                periods=re.findall(r'(?<![A-Za-z0-9])('+'|'.join(MONTHS)+r")[’' _-]?(\d{2})(?!\d)",text,re.I)
                if periods:break
            period=one([m.lower()+y for m,y in periods],'เดือนรายงาน')
            if period:result.update(month_name=MONTHS[period[:3]],year=period[3:])
        if task=='payment_online' and not result.get('mail_date') and selected:
            stamp=one([x['file_date'] for x in selected if x.get('file_date')],'วันที่ไฟล์') or received_date
            if stamp:result['mail_date']=datetime.strptime(stamp,'%y%m%d').strftime('%d%m%Y')
        if task in ('payment_online','customer_documents') and not selected and not (result.get('round') and result.get('mail_date')):
            raise ValidationError('จับประเภทงานได้ แต่ยังไม่มีรอบและวันที่ที่ชัดเจน')
        if task=='payment_online' and (not result.get('round') or not result.get('mail_date')):raise ValidationError('ยังระบุรอบหรือวันที่ Payment ไม่ได้')
        if task=='payment_online' and not 1<=int(result['round'])<=12:raise ValidationError('รอบ Payment ต้องอยู่ระหว่าง 1–12')
        if task in ('performance','performance_success_rate') and not result.get('month_name'):raise ValidationError('ยังระบุเดือนรายงานไม่ได้')
        if task=='enforcement' and not selected:
            if not result.get('mail_date'):result['needs_file_identity']=True
            result['decision']='Reject' if re.search('reject|ไม่อนุมัติ',subject+' '+body,re.I) else 'Approval'
    except (ValidationError,ValueError) as error:
        result['trigger_error']=str(error)
    return result


def logical_name(name):
    name=canonical_name(name)
    name=re.sub(r'^\d{6}_','',name)
    name=re.sub(r'^(A1015L?)_A1015L?_',r'\1_',name,flags=re.I)
    return re.sub(r'\.(?:zip|rar|pdf|xlsm|xlsx)(?:\.(?:gpg|pgp))?$|\.(?:gpg|pgp)$','',name,flags=re.I).casefold()


def infer_sources(folder,match,task):
    from pathlib import Path
    candidates=[]
    for path in Path(folder).iterdir():
        if not path.is_file():continue
        parsed=names(path.name,task)
        if len(parsed)!=1:continue
        data=parsed[0]
        if task=='customer_documents' and (data.get('round'),data.get('mail_date'))!=(match.get('round'),match.get('mail_date')):continue
        if task=='enforcement' and (data.get('decision','').casefold(),data.get('mail_date'))!=(match.get('decision','').casefold(),match.get('mail_date')):continue
        candidates.append(path)
    logical={}
    for path in candidates:logical.setdefault(logical_name(path.name),[]).append(path)
    if len(logical)>1:raise ValidationError('มีหลายไฟล์ตรงคำอธิบายเมล ต้องมีชื่อไฟล์ที่เจาะจงกว่านี้')
    if not logical:return []
    options=next(iter(logical.values()))
    if len(options)>1:raise ValidationError('มีไฟล์ซ้ำหลายชุดตรงเมล ยังไม่เลือกแทนโดยอัตโนมัติ')
    return options
