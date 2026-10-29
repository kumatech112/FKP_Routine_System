"""Read only the visible Outlook timestamp above each PDF table."""
from datetime import datetime
from collections import Counter
import re
from .ocr import recognize
WEEKDAYS=('จ.','อ.','พ.','พฤ.','ศ.','ส.','อา.')
STAMP=re.compile(r'(?<![0-9])([0-9]{1,2})\s*/\s*([0-9]{1,2})\s*/\s*([0-9]{4})\s+([0-9]{1,2})\s*:\s*([0-9]{2})(?![0-9])')

def parse_timestamp(text):
    values=set()
    text=text.translate(str.maketrans('๐๑๒๓๔๕๖๗๘๙','0123456789'))
    for match in STAMP.finditer(text):
        day,month,year,hour,minute=map(int,match.groups())
        try:date=datetime(year-543 if year>=2400 else year,month,day,hour,minute)
        except ValueError:continue
        values.add(f'{WEEKDAYS[date.weekday()]} {day:02}/{month:02}/{year:04} {hour:02}:{minute:02}')
    return next(iter(values)) if len(values)==1 else ''

def timestamp_consensus(texts):
    dates=Counter();times=Counter()
    for text in texts:
        stamp=parse_timestamp(text)
        if stamp:dates[stamp.split(' ')[1]]+=1
        matches=set(re.findall(r'(?<![0-9:])([0-9]{1,2}):([0-9]{2})(?![0-9:])',text))
        valid={f'{int(h):02}:{m}' for h,m in matches if int(h)<24 and int(m)<60}
        if len(valid)==1:times[next(iter(valid))]+=1
    date=list(dates)  # Keep only a unique valid printed date; time needs agreement between views.
    time=[v for v,n in times.items() if n>=2]
    return parse_timestamp(date[0]+' '+time[0]) if len(date)==len(time)==1 else ''

def pdf_timestamp(image,grids,executable):
    import cv2
    import numpy as np
    from PIL import ImageOps
    if not grids:return ''
    top=min(grid['y'][0] for grid in grids)
    if top<30:return ''
    crop=image.crop((round(image.width*.55),max(0,top-round(image.height*.32)),image.width,top-4))
    gray=np.asarray(crop.convert('L'))
    ink=gray<200
    # Isolate horizontal text lines; the date must be read apart from Outlook buttons.
    rows=np.flatnonzero(ink.sum(axis=1)>5)
    bands=[]
    for row in rows:
        if not bands or row>bands[-1][-1]+4:bands.append([int(row)])
        else:bands[-1].append(int(row))
    candidates=[]
    for band in bands:
        y0,y1=band[0],band[-1]+1
        if not max(5,image.height*.002)<=y1-y0<=image.height*.02:continue
        xs=np.flatnonzero(ink[y0:y1].any(axis=0))
        if not len(xs):continue
        if xs[-1]-xs[0]<image.width*.04:continue
        line=crop.crop((max(0,int(xs[0])-5),max(0,y0-5),min(crop.width,int(xs[-1])+6),min(crop.height,y1+5)))
        text,_=recognize(ImageOps.expand(line,12,'white'),executable,psm=13,lang='eng')
        stamp=parse_timestamp(text)
        if stamp:candidates.append((line,text,stamp))
    if not candidates or len({c[2] for c in candidates})!=1:return ''
    line,text,_=candidates[0]
    second,_=recognize(line,executable,psm=7,lang='eng')
    context,_=recognize(crop,executable,psm=11,lang='eng')
    return timestamp_consensus([text,second,context])
