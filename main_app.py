import calendar
import datetime as dt
import hashlib
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st
from core import SEOUL, load_workbook, popular_chart, search, fetch_videos, video_id, topic_stats, unique_videos
from trend_view import render as render_trends
from config_keys import youtube_api_key
import export_generic
import template_report

st.set_page_config(page_title='유튜브 노밤티 분석 보고서',page_icon='📊',layout='wide')
st.markdown('''<style>
.stApp,[data-testid="stAppViewContainer"],[data-testid="stHeader"],[data-testid="stSidebar"]{background:#fff;color:#111}
h1,h2,h3,p,label,[data-testid="stMetricValue"],[data-testid="stMetricLabel"]{color:#111}
[data-testid="stMetric"]{background:#f5f8fc;border:1px solid #cbd5e1;border-radius:10px;padding:16px}
[data-testid="stTabs"] [role="tablist"]{gap:5px;border-bottom:2px solid #172d4d}
[data-testid="stTabs"] [role="tab"]{background:#edf2f8;color:#172d4d;padding:10px;border-radius:5px 5px 0 0}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{background:#172d4d;color:white}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p{color:white}
input,textarea{background:#fff!important;color:#111!important}
[data-baseweb="input"],[data-baseweb="textarea"]{border:1px solid #94a3b8;border-radius:6px}
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] p{color:#111!important}
</style>''',unsafe_allow_html=True)
st.title('유튜브 노밤티 분석 보고서')
st.caption('엑셀 콘텐츠 성과와 공개 YouTube 데이터를 확인합니다.')
with st.sidebar:
    st.header('데이터 설정')
    upload=st.file_uploader('엑셀 파일 업로드',type=['xlsx'])
    default=None
    st.caption('배포본에는 원본 엑셀이 포함되지 않습니다. 분석할 파일을 업로드해 주세요.')
    try:
        try:
            configured_key=str(st.secrets.get('YOUTUBE_API_KEY','')).strip()
        except Exception:
            configured_key=''
        configured_key=configured_key or youtube_api_key()
    except (OSError,UnicodeError):
        configured_key=''
        st.warning('API 설정을 읽을 수 없습니다. Streamlit Secrets를 확인해 주세요.')
    entered_key=st.text_input('YouTube Data API 키 (직접 입력 시 우선)',type='password',help='직접 입력한 키는 세션에서만 사용됩니다.').strip()
    key=entered_key or configured_key
    if configured_key and not entered_key: st.caption('저장된 API 키를 사용합니다. 값은 화면에 표시하지 않습니다.')
    if st.button('API 연결 확인'):
        if not key: st.warning('먼저 API 키를 입력해 주세요.')
        else:
            try:
                popular_chart(key,limit=1)
                st.success('YouTube API 연결이 정상입니다.')
            except ValueError as e: st.error(str(e))
    st.caption('API 키 없이 엑셀 분석을 이용할 수 있습니다. 내부 유입 검색어·시청 유지율 자동 조회는 채널 관리자 인증이 필요합니다.')

df=pd.DataFrame(); source=''; detail_source=None
try:
    if upload:
        source=upload.name;df,issues,sheets=load_workbook(upload);upload.seek(0);detail_source=upload
    elif default:
        source=default.name;df,issues,sheets=load_workbook(default);detail_source=default
    if not df.empty:
        with st.expander(f'엑셀 확인 사항 ({len(issues)}건)'):
            for issue in issues:st.warning(issue)
except Exception as e:st.error(f'엑셀을 읽을 수 없습니다: {e}')

today=dt.datetime.now(SEOUL).date()
for name in ('weekly_results','million_results','chart_results','detail_video'):
    if name not in st.session_state:st.session_state[name]=None

def fmt(x):
    if x is None or pd.isna(x):return '기록 없음'
    return f'{x:,.0f}' if isinstance(x,(float,int)) else str(x)

def downloads(label,meta,columns,rows,identifier):
    st.caption('보고서는 현재 화면의 분석 결과와 기록된 출처를 포함합니다.')
    types={'PNG':('image/png','png',export_generic.png),'PDF':('application/pdf','pdf',export_generic.pdf),'Word':('application/vnd.openxmlformats-officedocument.wordprocessingml.document','docx',export_generic.docx)}
    choice=st.selectbox('출력 형식',list(types),key='type_'+identifier)
    try:
        mime,ext,build=types[choice]
        payload=build(label,meta,columns,rows)
        st.download_button(f'{choice} 다운로드',payload,file_name=f'{identifier}.{ext}',mime=mime,key='download_'+identifier)
        if choice=='PNG':
            with st.expander('PNG 미리보기'):st.image(payload,use_container_width=True)
    except Exception as e:st.error(f'{choice} 생성 오류: {e}')

def time_now():return dt.datetime.now(SEOUL).strftime('%Y-%m-%d %H:%M:%S KST')
def meta(source_text,period,extra=''):
    return [f'분석 기간: {period}',f'출처: {source_text}',f'작성 시점: {time_now()}',extra]

def score_summary(part):
    valid=part.dropna(subset=['조회수']).sort_values('조회수',ascending=False)
    if len(valid)<2:return '비교 가능한 조회수 기록이 부족합니다.'
    return f"누적 조회수 상위: {valid.iloc[0]['제목']} ({fmt(valid.iloc[0]['조회수'])}회), 하위: {valid.iloc[-1]['제목']} ({fmt(valid.iloc[-1]['조회수'])}회). 집계 시점·광고비·경과일의 차이를 함께 확인하세요."

def excel_report(start,end,label):
    if df.empty:st.info('엑셀 파일을 선택해 주세요.');return
    part=df[(df['발행일'].dt.date>=start)&(df['발행일'].dt.date<=end)].copy()
    if part.empty:st.info('이 기간에 발행된 영상이 없습니다.');return
    views=part['조회수'].dropna(); cost=part['광고비'].dropna(); likes=part['좋아요'].dropna(); comments=part['댓글수'].dropna()
    a,b,c=st.columns(3);a.metric('발행 영상',len(part));b.metric('엑셀 기록 누적 조회수',fmt(views.sum())+'회' if len(views) else '기록 없음');c.metric('기록된 광고비',fmt(cost.sum())+'원' if len(cost) else '기록 없음')
    st.caption(f'조회수 {len(views)}/{len(part)}건, 광고비 {len(cost)}/{len(part)}건 기록. 누적 조회수는 선택 기간 발생 조회수가 아닙니다.')
    st.write(f'기록된 좋아요 합계: {fmt(likes.sum()) if len(likes) else "기록 없음"} | 댓글 합계: {fmt(comments.sum()) if len(comments) else "기록 없음"}')
    st.dataframe(part[['발행일','제목','소재','조회수','좋아요','댓글수','광고비','광고조회수','CPV','평균조회율']],hide_index=True,use_container_width=True)
    summary=score_summary(part);st.info(summary)
    groups=part.groupby('소재',dropna=False).agg(영상수=('제목','count'),누적조회수_평균=('조회수','mean'),조회수_기록수=('조회수','count')).reset_index()
    st.subheader('소재별 기록');st.dataframe(groups,hide_index=True,use_container_width=True)
    st.write('운영 제안: 조회수가 높은 소재의 내용과 낮은 소재의 제목·대상을 함께 검토하고, 광고 집행 및 발행 후 경과 기간을 맞춰 비교하세요.')
    m=meta(f'{source} / 1. 신규',f'{start} ~ {end}',f'영상 {len(part)}건; 조회수 기록 {len(views)}건; 광고비 기록 {len(cost)}건. 누적 조회수이며 기간 조회수가 아닙니다.')
    m += [summary,'제안: 소재별 영상과 제목을 검토하되 광고비 및 발행 경과 기간 차이를 함께 확인하세요.',f'기록된 좋아요 {fmt(likes.sum()) if len(likes) else "기록 없음"}, 댓글 {fmt(comments.sum()) if len(comments) else "기록 없음"}.']
    records=[[r['발행일'].strftime('%Y-%m-%d'),r['제목'],fmt(r['조회수']),fmt(r['광고비'])] for _,r in part.sort_values('발행일',ascending=False).iterrows()]
    choice=st.selectbox('출력 형식',['PNG','PDF','Word'],key='template_type_'+label)
    formats={'PNG':('image/png','png',template_report.png),'PDF':('application/pdf','pdf',template_report.pdf),'Word':('application/vnd.openxmlformats-officedocument.wordprocessingml.document','docx',template_report.docx)}
    try:
        mime,extension,build=formats[choice]
        payload=build(part,start,end,label,source)
        st.download_button(f'{choice} 다운로드',payload,file_name=f'{label}_report.{extension}',mime=mime,key='template_download_'+label)
        if choice=='PNG':
            with st.expander('PNG 미리보기'):st.image(payload,use_container_width=True)
        if choice=='Word':st.caption('템플릿의 디자인을 유지하기 위해 Word 보고서는 페이지 이미지로 구성됩니다.')
    except Exception as e:st.error(f'{choice} 생성 오류: {e}')

def public_list(rows):
    for r in rows:
        with st.container(border=True):
            a,b=st.columns([1,4])
            with a:
                if r.get('썸네일'):st.image(r['썸네일'],use_container_width=True)
            with b:
                st.write(r['제목']);st.caption(f"{r.get('채널','')} | 발행 {str(r.get('발행일',''))[:10]} | 누적 조회수 {fmt(r.get('조회수'))} | 좋아요 {fmt(r.get('좋아요'))} | 댓글 {fmt(r.get('댓글수'))}"+(f" | 분류 {r['분류']} (ID {r['분류 ID']})" if r.get('분류 ID') and r.get('분류') else (f" | 분류 ID {r['분류 ID']}" if r.get('분류 ID') else '')))
                st.link_button('영상 보기',r['URL'])
                if r.get('설명'):
                    with st.expander('설명 보기'):st.write(r['설명'])

def keyword_display(rows,title,source_text,period,identifier):
    stats=topic_stats(rows)
    if not stats:st.info('두 영상 이상에서 반복된 단어가 없습니다.');return []
    st.caption(f'영상 {len(rows)}건의 제목·설명에서 추출한 단어입니다. 각 영상에서 한 번만 계산하며 형태소 분석은 하지 않습니다.')
    st.dataframe(pd.DataFrame(stats[:40]),hide_index=True,use_container_width=True)
    reportrows=[[x['키워드'],str(x['등장 영상 수']),x['예시 제목'],x['예시 URL']] for x in stats[:40]]
    downloads(title,meta(source_text,period,f'표본 영상 {len(rows)}건. 실제 검색량이나 유입 검색어가 아닙니다.'),['키워드','영상 수','예시 영상','예시 URL'],reportrows,identifier)
    return stats

def video_report(rows,title,period,at,identifier):
    if not rows:return
    data=[[str(r.get('발행일',''))[:10],r.get('제목',''),fmt(r.get('조회수')),r.get('URL','')] for r in rows]
    downloads(title,meta('YouTube Data API 공개 데이터 / 조회 '+at,period,f'표본 {len(rows)}건. 조회수는 조회 시점의 누적값입니다. 전체 목록이나 실제 유입 검색어가 아닙니다.'),['발행일','콘텐츠','누적 조회수','영상 URL'],data,identifier)

def require_key():
    if not key:st.warning('왼쪽에서 YouTube Data API 키를 입력해 주세요.');return False
    return True

weekly,monthly,detail,trends,million=st.tabs(['주간 리포트','월간 리포트','콘텐츠 상세','주간·월간 트렌드','100만 조회 콘텐츠'])
with weekly:
    st.subheader('주간 리포트');a,b=st.columns(2);start=a.date_input('시작일',today-dt.timedelta(days=6),key='week_start');end=b.date_input('종료일',today,key='week_end')
    if end<start:st.error('종료일이 시작일보다 빠릅니다.')
    else:excel_report(start,end,'주간')
with monthly:
    st.subheader('월간 리포트')
    if df.empty:st.info('엑셀 파일을 선택해 주세요.')
    else:
        months=sorted(df['발행일'].dt.strftime('%Y-%m').unique(),reverse=True);sel=st.selectbox('발행 월',months);y,m=map(int,sel.split('-'))
        excel_report(dt.date(y,m,1),dt.date(y,m,calendar.monthrange(y,m)[1]),'월간')
with detail:
    st.subheader('콘텐츠 상세')
    if df.empty:st.info('엑셀 파일을 선택해 주세요.')
    else:
        query=st.text_input('제목 검색');opts=df[df['제목'].fillna('').astype(str).str.contains(query,case=False,regex=False)].sort_values('발행일',ascending=False)
        if not opts.empty:
            labels={int(r['행']):f"{r['발행일']:%Y-%m-%d} | {r['제목']} (엑셀 {r['행']}행)" for _,r in opts.iterrows()};idx=st.selectbox('콘텐츠 선택',list(labels),format_func=lambda n,labels=labels:labels[n])
            try:
                if upload:upload.seek(0)
                sheet=openpyxl.load_workbook(detail_source,data_only=True,read_only=True)['1. 신규']
                names=['번호','발행일','재생목록','제목','소재','조회수','광고비','광고 집행 기간','광고 조회수','평균 CPV','좋아요','댓글','남성','여성','18~24세','25~34세','35~44세','45세 이상','구독자 증감','YouTube 광고','YouTube 검색','추천 동영상','탐색 기능','Shorts 피드','채널 페이지','노출 수','노출 클릭률','영상 길이','평균 시청 지속 시간','평균 조회율']
                vals=[sheet.cell(idx,col).value for col in range(2,32)]; percentages=set(range(12,18))|set(range(19,25))|{26,29}
                data=[]
                for i,(name,value) in enumerate(zip(names,vals)):
                    if value is None or value=='-': value='기록 없음'
                    elif i in percentages and isinstance(value,(float,int)):value=f'{value*100:.1f}%'
                    data.append([name,str(value)])
                st.dataframe(pd.DataFrame(data,columns=['항목','엑셀 기록']),hide_index=True,use_container_width=True)
                downloads('콘텐츠 상세',meta(source,'엑셀 기록 기준','원본 기록이며 측정 시점이 다를 수 있습니다.'),['항목','값'],data,'detail')
            except Exception as e:st.error(f'상세 정보를 읽지 못했습니다: {e}')
        else:st.info('콘텐츠를 찾지 못했습니다.')
    url=st.text_input('공개 영상 URL 또는 ID',key='detail_url')
    if st.button('공개 영상 조회') and require_key():
        st.session_state.detail_video=None
        try:
            rows=fetch_videos(key,[video_id(url)]);st.session_state.detail_video={'rows':rows,'at':time_now()}
            if not rows:st.warning('영상을 찾지 못했습니다.')
        except ValueError as e:st.error(str(e))
    if st.session_state.detail_video and st.session_state.detail_video['rows']:
        v=st.session_state.detail_video['rows'][0];st.caption('YouTube 공개 데이터 조회: '+st.session_state.detail_video['at']);public_list([v])
        if not df.empty:
            direct=df[df['영상ID']==v['영상ID']]
            if len(direct):st.success('영상 ID가 일치하는 엑셀 행: '+', '.join(map(str,direct['행'])))
            else:
                similar=df[df['제목'].fillna('').map(lambda s:SequenceMatcher(None,str(s).lower(),v['제목'].lower()).ratio()>=.65)]
                if len(similar):st.warning('제목 유사 후보입니다. 동일 영상인지 직접 확인하세요.');st.dataframe(similar[['행','제목','발행일']],hide_index=True)
with trends:
    render_trends(key,today,time_now,keyword_display,video_report)
with million:
    st.subheader('100만 조회 콘텐츠 상위 10개')
    st.info('검색어 없이 한국에서 시청 가능한 공개 영상 표본 최대 50개를 조회해, 누적 조회수 100만 회 이상인 영상 중 상위 10개를 표시합니다. 기본 범위는 최근 30일에 발행된 영상입니다.')
    st.caption('API가 반환한 일부 검색 결과의 순위이며 YouTube 전체의 공식 순위가 아닙니다. 조회수는 조회 시점의 누적값으로 광고 효과가 포함될 수 있습니다.')
    a,b=st.columns(2)
    start=a.date_input('발행 시작일',today-dt.timedelta(days=29),key='million_start')
    end=b.date_input('발행 종료일',today,key='million_end')
    refresh=st.button('상위 콘텐츠 새로고침',key='million_refresh')
    if not key:
        st.session_state.million_results=None
        st.session_state.pop('million_request',None)
        st.info('왼쪽에 YouTube API 키를 입력하고 Enter를 누르면 자동으로 불러옵니다.')
    elif end<start:
        st.session_state.million_results=None
        st.error('발행 종료일은 시작일보다 빠를 수 없습니다.')
    else:
        fingerprint=hashlib.sha256((key+'\0'+str(start)+'\0'+str(end)).encode()).hexdigest()
        if refresh or st.session_state.get('million_request')!=fingerprint:
            st.session_state['million_request']=fingerprint
            st.session_state.million_results=None
            st.session_state['million_error']=None
            try:
                with st.spinner('조회수 상위 영상을 불러오고 있습니다...'):
                    ids=search(key,'',start,end,50,order='viewCount')
                    found=unique_videos(fetch_videos(key,ids))
                eligible=sorted((x for x in found if x['조회수'] is not None and x['조회수']>=1000000),key=lambda x:x['조회수'],reverse=True)
                st.session_state.million_results={'rows':eligible[:10],'period':f'{start} ~ {end}','at':time_now(),'sample':len(found),'eligible':len(eligible)}
            except ValueError as e:st.session_state['million_error']=str(e)
        if st.session_state.get('million_error'): st.error(st.session_state['million_error'])
    data=st.session_state.million_results if key and end>=start else None
    if data:
        st.caption(f"발행 기간 {data['period']} | 조회 {data['at']} | 검색 표본 {data['sample']}건 | 100만 회 이상 {data['eligible']}건 | 표시 {len(data['rows'])}건")
        if data['rows']:
            for rank,row in enumerate(data['rows'],1):
                st.markdown(f'#### {rank}위')
                public_list([row])
            keyword_display(data['rows'],'조회수 기준 콘텐츠 키워드',f"YouTube 공개 검색 / 100만 회 이상 상위 10개 / {data['at']}",data['period'],'million')
            video_report(data['rows'],'조회수 기준 콘텐츠',data['period'],data['at'],'million_videos')
        else:st.info('조회된 표본에서 누적 조회수 100만 회 이상인 영상이 없습니다. 발행 기간을 넓혀 다시 확인해 주세요.')
