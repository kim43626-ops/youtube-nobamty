from io import BytesIO
from html import escape
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from fonts import FONT, BOLD

def clean(value):
    if value is None: return '기록 없음'
    text=str(value)
    return '기록 없음' if text in ('nan','NaT','None','') else text

def wrap(draw,text,font,width):
    lines=[]; line=''
    for ch in clean(text):
        if ch=='\n': lines.append(line); line=''; continue
        if draw.textlength(line+ch,font=font)>width and line:
            lines.append(line); line=''
        line+=ch
    if line: lines.append(line)
    return lines or ['']

def png(title,meta,columns,rows):
    regular=ImageFont.truetype(FONT,22); bold=ImageFont.truetype(BOLD,24); heading=ImageFont.truetype(BOLD,38)
    probe=ImageDraw.Draw(Image.new('RGB',(1,1))); width=1600; margin=70; usable=width-2*margin
    cols=max(1,len(columns)); cw=usable/cols
    meta_lines=[v for line in meta for v in wrap(probe,line,regular,usable)]
    wrapped=[[wrap(probe,v,regular,cw-18) for v in row] for row in rows]
    heights=[max(len(c) for c in row)*30+20 for row in wrapped]
    height=240+len(meta_lines)*34+55+sum(heights)+90
    if height>60000: raise ValueError('보고서가 너무 큽니다. 기간 또는 표본을 줄여 주세요.')
    im=Image.new('RGB',(width,height),'white'); d=ImageDraw.Draw(im); navy='#111111'; dark='#111111'; border='#dedede'
    d.rectangle((0,0,width,175),fill=navy); d.text((margin,28),'YOUTUBE REPORT',font=regular,fill='#cccccc'); d.text((margin,85),title,font=heading,fill='white'); y=205
    for line in meta_lines: d.text((margin,y),line,font=regular,fill=dark); y+=34
    y+=20; d.rectangle((margin,y,width-margin,y+46),fill=navy)
    for i,col in enumerate(columns): d.text((margin+i*cw+8,y+7),str(col),font=bold,fill='white')
    y+=46
    for cells,h in zip(wrapped,heights):
        for i,lines in enumerate(cells):
            x=margin+i*cw; d.rectangle((x,y,x+cw,y+h),outline=border,width=1)
            for j,line in enumerate(lines): d.text((x+8,y+8+j*30),line,font=regular,fill=dark)
        y+=h
    out=BytesIO(); im.save(out,format='PNG'); return out.getvalue()

def pdf(title,meta,columns,rows):
    pdfmetrics.registerFont(TTFont('KoreanReport',FONT))
    out=BytesIO(); doc=SimpleDocTemplate(out,pagesize=A4,leftMargin=36,rightMargin=36,topMargin=36,bottomMargin=36)
    body=ParagraphStyle('body',fontName='KoreanReport',fontSize=8,leading=12,wordWrap='CJK')
    head=ParagraphStyle('head',parent=body,fontSize=17,leading=24,textColor=colors.white)
    P=lambda s:Paragraph(escape(clean(s)).replace('\n','<br/>'),body)
    cover=LongTable([[Paragraph(escape(title),head)]],colWidths=[A4[0]-72])
    cover.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#111111')),('TOPPADDING',(0,0),(-1,-1),22),('BOTTOMPADDING',(0,0),(-1,-1),22),('LEFTPADDING',(0,0),(-1,-1),16)]))
    story=[cover,Spacer(1,12)]+[P(x) for x in meta]+[Spacer(1,10)]
    data=[[P(v) for v in columns]]+[[P(v) for v in row] for row in rows]
    table=LongTable(data,colWidths=[(A4[0]-72)/len(columns)]*len(columns),repeatRows=1)
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('LINEABOVE',(0,0),(-1,0),1,colors.HexColor('#111111')),('LINEBELOW',(0,0),(-1,0),.7,colors.HexColor('#111111')),('LINEBELOW',(0,1),(-1,-1),.3,colors.HexColor('#dddddd')),('VALIGN',(0,0),(-1,-1),'TOP'),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    story.append(table);doc.build(story);return out.getvalue()

def docx(title,meta,columns,rows):
    doc=Document(); style=doc.styles['Normal'];style.font.name='맑은 고딕';style._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),'맑은 고딕');style.font.size=Pt(9)
    section=doc.sections[0];section.left_margin=Cm(1.8);section.right_margin=Cm(1.8)
    cover=doc.add_table(rows=1,cols=1);cell=cover.cell(0,0);cell.text='';cell._tc.get_or_add_tcPr().append(OxmlElement('w:shd'));cell._tc.tcPr[-1].set(qn('w:fill'),'111111')
    run=cell.paragraphs[0].add_run(title);run.bold=True;run.font.size=Pt(19);run.font.color.rgb=RGBColor(255,255,255)
    for line in meta:doc.add_paragraph(clean(line))
    table=doc.add_table(rows=1,cols=len(columns));table.style='Table Grid'
    for cell,value in zip(table.rows[0].cells,columns):cell.text=clean(value)
    for row in rows:
        for cell,value in zip(table.add_row().cells,row):cell.text=clean(value)
    out=BytesIO();doc.save(out);return out.getvalue()
