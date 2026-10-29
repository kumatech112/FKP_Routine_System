"""High-resolution table OCR with an optional workflow-controlled review gate."""
from pathlib import Path
from statistics import mean
import json
import re

from .pdf_timestamp import pdf_timestamp
from .vendor.fkp.domain import ValidationError
from .vendor.fkp.persistence import digest, dump, publish_outputs

FIELDS = (
    "ลำดับ", "บังคับวิชาชีพ", "OA", "Product", "acc", "ชื่อ-นามสกุล",
    "วันที่ส่งเรื่องขออนุมัติยึด", "Bill Code ก่อนขออนุมัติยึด",
    "ประเภทอสังหาริมทรัพย์ที่จะยึด", "วันที่หมายบังคับคดีออก", "ราคาประเมินทรัพย์",
    "ประเภทการจำนอง", "ราคาจำนอง", "วันที่คำนวณยอดหนี้ตามคำพิพากษา",
    "ยอดหนี้ตามคำพิพากษา", "ชื่อศาล", "ข้อมูลโฉนดที่ดินที่จะยึด",
    "ความเห็น R&L ผู้รับผิดชอบ", "ความเห็น R&L",
    "อนุมัติ/ไม่อนุมัติ R&L Mgr.Up", "ความเห็น R&L (ถ้ามี) R&L Mgr.Up",
)


def centers(values, gap=3):
    groups = []
    for value in values:
        if not groups or value > groups[-1][-1] + gap:
            groups.append([int(value)])
        else:
            groups[-1].append(int(value))
    return [int(round(mean(group))) for group in groups]


def table_grids(image, previous=None):
    """Find header tables and continued rows, retaining clipped row fragments."""
    import cv2
    import numpy as np
    rgb = np.asarray(image.convert("RGB"))
    mask = (rgb.max(axis=2) < 105).astype("uint8") * 255
    height, width = mask.shape
    horizontal = cv2.morphologyEx(mask, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, width//40), 1)))
    vertical = cv2.morphologyEx(mask, cv2.MORPH_OPEN,
        cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(25, height//80))))
    contours, _ = cv2.findContours(horizontal | vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    found = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w < width*.35 or h < height*.02:
            continue
        rows = centers(np.flatnonzero((horizontal[y:y+h, x:x+w] > 0).sum(axis=1) > w*.55))
        if not rows:
            continue
        top_gap = rows[0] > 10
        prior = previous and abs(x/width-previous["x"][0]/previous["width"]) < .025
        if prior:
            prior = abs((x+w-1)/width-previous["x"][-1]/previous["width"]) < .025
        if len(rows) >= 2:
            band_rgb = rgb[y+rows[0]+3:y+rows[1]-3, x:x+w]
            fill = ((band_rgb.max(axis=2) > 105) & (band_rgb.min(axis=2) < 230)).mean() if band_rgb.size else 0
        else:
            fill = 0
        continued = bool(prior and (top_gap or fill < .15))
        if continued:
            columns = [round(c*width/previous["width"])-x for c in previous["x"]]
            band = vertical[y:y+h, x:x+w]
            actual = centers(np.flatnonzero((band > 0).sum(axis=0) > h*.5))
            if len(actual) < max(4, len(columns)-3):
                continue
            if top_gap:
                rows.insert(0, 0)
        else:
            if len(rows) < 3:
                continue
            top, bottom = rows[:2]
            band = vertical[y+top:y+bottom+1, x:x+w]
            columns = centers(np.flatnonzero((band > 0).sum(axis=0) > (bottom-top)*.65))
            if len(columns) < 4:
                continue
            if w-1-columns[-1] > max(12, (columns[-1]-columns[-2])*.25):
                columns.append(w-1)
        partial_bottom = h-1-rows[-1] > 12
        if partial_bottom:
            rows.append(h-1)
        if len(rows) < (2 if continued else 3):
            continue
        found.append({"x":[x+c for c in columns], "y":[y+r for r in rows],
            "width":width, "continued":continued, "partial_top":continued and top_gap,
            "partial_bottom":partial_bottom})
    return sorted(found, key=lambda table:table["y"][0])


from .ocr import configure_ocr,recognize


def ocr_image(image, executable, psm=6, header=False, field=''):
    import cv2
    import numpy as np
    from PIL import Image,ImageOps
    if psm!=6:
        return recognize(image,executable,psm=psm)
    interior=image.crop((12,12,image.width-12,image.height-12))
    rgb=np.asarray(interior.convert('RGB')).astype('int16')
    if header:
        background=np.median(rgb.reshape(-1,3),axis=0)
        contrast=np.max(np.abs(rgb-background),axis=2).astype('uint8')
        _,foreground=cv2.threshold(contrast,32,255,cv2.THRESH_BINARY)
        interior=Image.fromarray(255-foreground)
    else:
        # Excel's green top-left markers are graphics, not account-number glyphs.
        green=(rgb[:,:,1]>rgb[:,:,0]+25)&(rgb[:,:,1]>rgb[:,:,2]+25)
        rgb[green]=255
        interior=Image.fromarray(rgb.astype('uint8'))
    key=re.sub(r'\s+','',field).casefold()
    digits=not header and (key in ('acc','oa','ลำดับ') or 'billcode' in key)
    english=not header and key=='product'
    language='eng' if digits or english else 'tha+eng'
    whitelist='0123456789' if digits else ''
    comment=not header and ('ความเห็น' in key or 'mgr.up' in key)
    scales=(1,) if digits or english else ((2,1.6,1) if comment else (1.6,2,1))
    candidates=[]
    for scale in scales:
        resized=interior.resize((max(1,round(interior.width*scale)),interior.height),Image.Resampling.LANCZOS)
        text,score=recognize(ImageOps.expand(resized,12,'white'),executable,lang=language,whitelist=whitelist)
        candidates.append((text,score))
        if score>=85:break
    return max(candidates,key=lambda pair:pair[1])


def image_mail_text(images, executable, job):
    from PIL import Image
    result = []
    for path in images:
        job.check()
        with Image.open(path) as image:
            if image.width < 180 or image.height < 35:
                continue
            image = image.convert("RGB")
            if image.width < 2000:
                scale = min(3, 2000 / image.width)
                image = image.resize((int(image.width*scale), int(image.height*scale)))
            text, score = ocr_image(image, executable, psm=11)
            result.append(text)
    text = "\n".join(result)
    # File names may wrap across image lines. Only strict A1015 patterns are used later.
    compact = re.sub(r"\s+", "", text)
    dump(job.folder / "mail_image_ocr.json", {"text": text, "compact": compact})
    return text + "\n" + compact


def prepare_enforcement(folder, job, executable):
    """OCR only table regions; join fragments when a row crosses PDF pages."""
    import pymupdf
    from PIL import Image, ImageOps
    configure_ocr(executable)
    pdfs = sorted(p for p in Path(folder).rglob("*") if p.is_file() and p.suffix.casefold()==".pdf")
    if not pdfs:
        raise ValidationError("ไม่พบ PDF ในไฟล์ Enforcement")
    review = {"version":3, "status":"needs_review", "cells":[], "pages":[], "tables":[],
              "pdfs":[{"path":str(p),"sha256":digest(p)} for p in pdfs],
              "output":str(job.folder/"Enforcement.xlsx")}
    for file_number,pdf in enumerate(pdfs,1):
        previous = None
        record_number = 0
        file_stamps=set()
        with pymupdf.open(pdf) as document:
            if document.is_encrypted:
                raise ValidationError("PDF ยังเข้ารหัส")
            for page_number,page in enumerate(document,1):
                job.check()
                pixmap=page.get_pixmap(dpi=400,alpha=False)
                image=Image.frombytes("RGB",[pixmap.width,pixmap.height],pixmap.samples)
                page_id=f"{file_number}:{page_number}"
                page_path=job.folder/"OCR"/f"file_{file_number}_page_{page_number}.png"
                page_path.parent.mkdir(parents=True,exist_ok=True)
                image.save(page_path)
                grids=table_grids(image,previous)
                page_stamp=pdf_timestamp(image,grids,executable)
                if page_stamp:file_stamps.add(page_stamp)
                review["pages"].append({"id":page_id,"source":str(pdf),"page":page_number,
                    "image":str(page_path),"table_count":len(grids),"pdf_datetime":page_stamp})
                for table_number,grid in enumerate(grids,1):
                    xs,ys=grid["x"],grid["y"]
                    table_id=f"{page_id}:{table_number}"
                    continued=grid["continued"] and previous is not None
                    if continued:
                        group_id=previous["group_id"]
                        headers=previous["headers"]
                    else:
                        group_id=table_id
                        headers=[]
                        for col in range(len(xs)-1):
                            job.check()
                            crop=ImageOps.expand(image.crop((xs[col]+3,ys[0]+3,xs[col+1]-3,ys[1]-3)),border=12,fill="white")
                            text,score=ocr_image(crop,executable,header=True)
                            text=text.strip() or f"ช่อง {col+1}"
                            headers.append(text)
                            crop_path=page_path.with_name(f"{page_path.stem}_t{table_number}_header_{col+1}.png")
                            crop.save(crop_path)
                            review["cells"].append({"id":f"{table_id}:header:{col}","group_id":group_id,
                                "table_id":table_id,"page_id":page_id,"row":0,"column":-2,"header_column":col,
                                "field":f"หัวตาราง {col+1}","value":text,"confidence":score,
                                "state":"needs_review","image":str(crop_path)})
                    first=0 if continued else 1
                    rows=len(ys)-1-first
                    review["tables"].append({"id":table_id,"group_id":group_id,"page_id":page_id,
                        "file_number":file_number,"headers":headers,"rows":rows,"continued":continued,
                        "partial_bottom":grid["partial_bottom"],"pdf_datetime":page_stamp})
                    last_record=None
                    for row_index in range(first,len(ys)-1):
                        if (continued and row_index==0 and grid["partial_top"]
                                and previous["partial_bottom"]):
                            record_id=previous["last_record"]
                        else:
                            record_number+=1
                            record_id=f"{file_number}:row:{record_number}"
                        last_record=record_id
                        for col in range(len(xs)-1):
                            job.check()
                            box=(xs[col]+3,ys[row_index]+3,xs[col+1]-3,ys[row_index+1]-3)
                            if box[2]<=box[0] or box[3]<=box[1]:
                                raise ValidationError("ขอบเขตช่องตารางไม่ถูกต้อง")
                            crop=ImageOps.expand(image.crop(box),border=12,fill="white")
                            text,score=ocr_image(crop,executable,field=headers[col])
                            crop_path=page_path.with_name(f"{page_path.stem}_t{table_number}_r{row_index}_c{col+1}.png")
                            crop.save(crop_path)
                            review["cells"].append({"id":f"{table_id}:{row_index}:{col}","group_id":group_id,
                                "table_id":table_id,"page_id":page_id,"record_id":record_id,
                                "row":record_number,"column":col,"field":headers[col],"value":text,
                                "confidence":score,"state":"needs_review","image":str(crop_path),"box":box,
                                "file_number":file_number,"pdf_datetime":page_stamp})
                    previous={**grid,"group_id":group_id,"headers":headers,"last_record":last_record}
                job.progress(10+70*(file_number-1+page_number/len(document))/len(pdfs),
                    f"OCR ตาราง PDF {file_number}/{len(pdfs)} หน้า {page_number}")
        if len(file_stamps)==1:
            stamp=next(iter(file_stamps))
            for cell in review['cells']:
                if cell.get('file_number')==file_number and not cell.get('pdf_datetime'):
                    cell['pdf_datetime']=stamp
        if not file_stamps:
            job.log(f'ไม่พบวันที่และเวลาบริเวณหัวอีเมลใน PDF: {pdf.name}; เว้นคอลัมน์วันที่ไว้')
        if digest(pdf)!=review["pdfs"][file_number-1]["sha256"]:
            raise ValidationError("PDF ต้นทางเปลี่ยนระหว่าง OCR")
    if not review["tables"]:
        raise ValidationError("ไม่พบตารางใน PDF กรุณาตรวจไฟล์ต้นทาง")
    review["page_count"]=len(review["pages"])
    review["cell_count"]=sum(c["column"]>=0 for c in review["cells"])
    dump(job.folder/"enforcement_review.json",review)
    write_workbook(review)
    return review


def write_workbook(review):
    """Export table cells only: no mail text, signatures, evidence or audit sheets."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment,Font,PatternFill
    from openpyxl.utils import get_column_letter
    book=Workbook()
    book.remove(book.active)
    groups={}
    for table in review["tables"]:
        group=table.get("group_id",table["id"])
        groups.setdefault(group,table)
    for group,table in groups.items():
        columns=len(table["headers"])
        headers=list(table["headers"])
        for cell in review["cells"]:
            if cell.get("group_id")==group and cell["column"]==-2 and "header_column" in cell:
                headers[cell["header_column"]]=cell["value"]
        number=len(book.worksheets)+1
        sheet=book.create_sheet("ตาราง" if number==1 else f"ตาราง {number}")
        include_stamp=review.get("version",2)>=3
        sheet.append(headers+(["วันที่อีเมลใน PDF"] if include_stamp else []))
        records={}
        record_stamps={}
        for cell in review["cells"]:
            if cell["column"]<0:
                continue
            cell_group=cell.get("group_id",cell.get("table_id"))
            if cell_group!=group:
                continue
            record=cell.get("record_id",f'{cell.get("table_id")}:{cell["row"]}')
            stamp=cell.get("pdf_datetime","")
            if stamp:
                if record in record_stamps and record_stamps[record]!=stamp:
                    raise ValidationError("วันที่ PDF ของรายการต่อเนื่องไม่ตรงกัน")
                record_stamps[record]=stamp
            values=records.setdefault(record,[""]*columns)
            col=cell["column"]
            text=cell["value"]
            if text and values[col] and text!=values[col]:
                values[col]+="\n"+text
            elif text:
                values[col]=text
        for record,values in records.items():
            sheet.append(values+([record_stamps.get(record,"")] if include_stamp else []))
        sheet.freeze_panes="A2"
        sheet.auto_filter.ref=sheet.dimensions
        for row in sheet:
            for cell in row:
                cell.font=Font(name="Tahoma",size=11,bold=cell.row==1)
                cell.alignment=Alignment(vertical="top",wrap_text=True)
                if isinstance(cell.value,str):
                    if len(cell.value)>32767:
                        raise ValidationError("ข้อความตารางเกินขีดจำกัด Excel ต้องตรวจทานก่อนส่ง")
                    cell.data_type="s"
                    cell.number_format="@"
                if cell.row==1:
                    cell.fill=PatternFill("solid",fgColor="D9E2F3")
        for col in range(1,columns+1+int(include_stamp)):
            sheet.column_dimensions[get_column_letter(col)].width=25
        sheet.row_dimensions[1].height=60
        for row in range(2,sheet.max_row+1):
            sheet.row_dimensions[row].height=90
    if not book.worksheets:
        raise ValidationError("ไม่มีตารางสำหรับสร้าง Excel")
    path=Path(review["output"])
    book.save(path)
    return path


def prepare_review(review_path,edits,require_review=True):
    review_path=Path(review_path)
    review=json.loads(review_path.read_text(encoding="utf-8-sig"))
    if review["status"]=="success":
        raise ValidationError("งานนี้ยืนยันสำเร็จแล้ว")
    if not review["tables"]:
        raise ValidationError("ไม่พบตารางสำหรับตรวจทาน")
    for cell in review["cells"]:
        if cell["id"] in edits:
            cell.update(value=edits[cell["id"]],state="reviewed")
    dump(review_path,review)
    if require_review and any(c["state"]!="reviewed" for c in review["cells"]):
        raise ValidationError("ยังมีช่องหรือหัวตารางที่ยังไม่ได้ตรวจทาน")
    write_workbook(review)
    return review


def approve_review(review_path,edits,job,settings,additional_outputs=()):
    review=prepare_review(review_path,edits)
    outputs=publish_outputs([review["output"],*additional_outputs],settings,job)
    review.update(status="success",outputs=outputs)
    dump(review_path,review)
    return outputs
