"""Report layout based on the supplied youtube_monthly_report_template.html."""
from io import BytesIO
import datetime as dt
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from docx import Document
from docx.shared import Cm, Pt

W,H=1240,1754
LEFT,RIGHT=95,1145
from fonts import FONT, BOLD
F=lambda size,bold=False:ImageFont.truetype(BOLD if bold else FONT,size)

def split_text(draw,text,font,width):
    out=[]; line=''
    for ch in str(text):
        if ch=='\n': out.append(line);line='';continue
        if line and draw.textlength(line+ch,font=font)>width:out.append(line);line=''
        line+=ch
    if line:out.append(line)
    return out or ['']

def text_lines(draw,position,text,font,width,fill='#111111',leading=32):
    x,y=position
    for line in split_text(draw,text,font,width):
        draw.text((x,y),line,font=font,fill=fill);y+=leading
    return y

def fmt(value):
    return '기록 없음' if pd.isna(value) else f'{float(value):,.0f}'

def pages(df,start,end,label,source):
    rows=df.sort_values('발행일',ascending=False).copy(); now=dt.datetime.now().strftime('%Y-%m-%d %H:%M')
    views=rows['조회수'].dropna();cost=rows['광고비'].dropna();likes=rows['좋아요'].dropna();comments=rows['댓글수'].dropna()
    main=Image.new('RGB',(W,H),'white'); d=ImageDraw.Draw(main)
    d.rectangle((0,0,W,315),fill='#111111');d.text((LEFT,58),'YOUTUBE REPORT',font=F(17,True),fill='#c8c8c8')
    d.text((LEFT,135),f'유튜브 {label} 리포트',font=F(53,True),fill='white');d.text((LEFT,235),f'{start:%Y-%m-%d} ~ {end:%Y-%m-%d}',font=F(24),fill='#dddddd')
    y=355;d.text((LEFT,y),'OVERVIEW',font=F(16,True),fill='#666666');y+=28;d.text((LEFT,y),'핵심 지표',font=F(31,True),fill='#111111');y+=64
    d.line((LEFT,y,RIGHT,y),fill='#111111',width=3); y+=15
    metrics=[('발행 영상',f'{len(rows):,}건'),('누적 조회수',fmt(views.sum())+'회' if len(views) else '기록 없음'),('광고비 합계',fmt(cost.sum())+'원' if len(cost) else '기록 없음'),('좋아요 / 댓글',(fmt(likes.sum()) if len(likes) else '기록 없음')+' / '+(fmt(comments.sum()) if len(comments) else '기록 없음'))]
    for i,(name,value) in enumerate(metrics):
        x=LEFT+i*(RIGHT-LEFT)/4
        if i:d.line((x-10,y,x-10,y+88),fill='#dddddd',width=1)
        d.text((x,y),name,font=F(18),fill='#666666')
        size=24
        while size>14 and d.textlength(value,font=F(size,True))>(RIGHT-LEFT)/4-18:size-=1
        d.text((x,y+32),value,font=F(size,True),fill='#111111')
    y+=105;d.line((LEFT,y,RIGHT,y),fill='#dedede',width=1);y+=35
    d.text((LEFT,y),'TOP PERFORMANCE',font=F(16,True),fill='#666666');y+=25;d.text((LEFT,y),'누적 조회수 상위 3개',font=F(31,True),fill='#111111');y+=55
    ranked=rows.dropna(subset=['조회수']).sort_values('조회수',ascending=False).head(3)
    maxview=ranked['조회수'].max() if len(ranked) else 0
    if ranked.empty:d.text((LEFT,y),'조회수 기록 없음',font=F(21),fill='#666666');y+=50
    for _,r in ranked.iterrows():
        name=str(r['제목'] or '제목 없음'); fit=split_text(d,name,F(18,True),RIGHT-LEFT)
        shown=fit[0] if len(fit)==1 else fit[0]+'…'
        d.text((LEFT,y),shown,font=F(18,True),fill='#111111');y+=30
        d.rectangle((LEFT,y,RIGHT,y+10),fill='#eeeeee')
        if maxview>0 and r['조회수']>0:d.rectangle((LEFT,y,LEFT+int((RIGHT-LEFT)*r['조회수']/maxview),y+10),fill='#111111')
        y+=16;d.text((LEFT,y),fmt(r['조회수'])+'회',font=F(17),fill='#666666');y+=38
    y+=12;d.text((LEFT,y),'INSIGHT',font=F(16,True),fill='#666666');y+=27;d.text((LEFT,y),'요약 및 제안',font=F(31,True),fill='#111111');y+=52
    lines=[f'조회수 기록 {len(views)}/{len(rows)}건, 광고비 기록 {len(cost)}/{len(rows)}건입니다.', '조회수는 엑셀 누적값입니다. 발행 경과 기간과 광고 집행 차이를 함께 확인하세요.']
    if len(ranked):lines.insert(0,f"누적 조회수 상위: {ranked.iloc[0]['제목']} ({fmt(ranked.iloc[0]['조회수'])}회)")
    block_lines=[]
    for line in lines:block_lines.extend(split_text(d,'• '+line,F(19),RIGHT-LEFT-50))
    height=len(block_lines)*30+28;d.rectangle((LEFT,y,RIGHT,y+height),fill='#f4f4f4');d.rectangle((LEFT,y,LEFT+5,y+height),fill='#111111')
    for i,line in enumerate(block_lines):d.text((LEFT+22,y+12+i*30),line,font=F(19),fill='#111111')
    y+=height+40
    d.text((LEFT,y),'DETAIL',font=F(16,True),fill='#666666');y+=27;d.text((LEFT,y),'콘텐츠 상세',font=F(31,True),fill='#111111');y+=56
    result=[]
    def header(draw,y):
        draw.line((LEFT,y,RIGHT,y),fill='#111111',width=3);y+=12
        for x,title in [(LEFT,'발행일'),(LEFT+150,'콘텐츠'),(LEFT+750,'누적 조회수'),(LEFT+910,'광고비')]:draw.text((x,y),title,font=F(17,True),fill='#555555')
        return y+40
    def footer(draw,page_num):
        draw.line((LEFT,H-115,RIGHT,H-115),fill='#dddddd',width=1)
        draw.text((LEFT,H-100),'조회수는 기간 조회수가 아닌 엑셀 기록 누적값입니다.',font=F(16),fill='#555555')
        draw.text((LEFT,H-72),'출처: '+str(source)[:70]+' / 1. 신규',font=F(14),fill='#666666')
        draw.text((LEFT,H-48),f'작성: {now}  |  {page_num}쪽',font=F(14),fill='#666666')
    y=header(d,y)
    for _,r in rows.iterrows():
        titles=split_text(d,r['제목'] or '제목 없음',F(18),580)
        h=max(52,len(titles)*27+18)
        if y+h>H-135:
            footer(d,len(result)+1);result.append(main)
            main=Image.new('RGB',(W,H),'white');d=ImageDraw.Draw(main)
            d.text((LEFT,55),f'유튜브 {label} 리포트 · 콘텐츠 상세',font=F(30,True),fill='#111111')
            y=header(d,120)
        d.text((LEFT,y+10),r['발행일'].strftime('%Y-%m-%d'),font=F(17),fill='#111111')
        for i,line in enumerate(titles):d.text((LEFT+150,y+8+i*27),line,font=F(18),fill='#111111')
        d.text((LEFT+750,y+10),fmt(r['조회수']),font=F(17),fill='#111111');d.text((LEFT+910,y+10),fmt(r['광고비']),font=F(17),fill='#111111')
        y+=h;d.line((LEFT,y,RIGHT,y),fill='#dedede',width=1)
    footer(d,len(result)+1);result.append(main)
    return result

def png(df,start,end,label,source):
    imgs=pages(df,start,end,label,source);combined=Image.new('RGB',(W,H*len(imgs)),'white')
    for i,im in enumerate(imgs):combined.paste(im,(0,i*H))
    out=BytesIO();combined.save(out,format='PNG');return out.getvalue()

def pdf(df,start,end,label,source):
    out=BytesIO();c=canvas.Canvas(out,pagesize=A4)
    for im in pages(df,start,end,label,source):c.drawImage(ImageReader(im),0,0,width=A4[0],height=A4[1]);c.showPage()
    c.save();return out.getvalue()

def docx(df,start,end,label,source):
    doc=Document();section=doc.sections[0];section.top_margin=Cm(1);section.bottom_margin=Cm(1);section.left_margin=Cm(1);section.right_margin=Cm(1)
    for i,im in enumerate(pages(df,start,end,label,source)):
        if i:doc.add_page_break()
        buf=BytesIO();im.save(buf,format='PNG');buf.seek(0)
        paragraph=doc.add_paragraph();paragraph.paragraph_format.space_after=Pt(0);paragraph.paragraph_format.space_before=Pt(0);paragraph.add_run().add_picture(buf,width=Cm(19))
    out=BytesIO();doc.save(out);return out.getvalue()
