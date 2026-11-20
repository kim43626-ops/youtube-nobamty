import calendar
import datetime as dt
import hashlib

import streamlit as st
from core import search, fetch_videos, topic_stats, unique_videos


def periods(day, week):
    first=dt.date(day.year,day.month,1)
    last=calendar.monthrange(day.year,day.month)[1]
    start_day=(week-1)*7+1
    if start_day>last: raise ValueError('선택한 월에는 해당 주차가 없습니다.')
    return (first.replace(day=start_day), first.replace(day=min(start_day+6,last))), (first,first.replace(day=last))


def render(key,today,time_now,keyword_display,video_report):
    st.subheader('주간·월간 트렌드')
    st.write('월과 주차를 선택하면 YouTube 공개 검색 결과에서 반복되는 키워드를 자동으로 정리합니다.')
    st.info('한국에서 시청 가능한 영상 중 입력한 주제에 대한 검색 표본입니다. 실제 검색량·전체 인기 순위·유입률이 아니며, 기간 중 발행된 영상의 현재 공개 정보입니다.')
    a,b=st.columns(2)
    month=a.date_input('분석할 월',value=today.replace(day=1),key='trend_month',help='날짜를 선택하면 해당 월 전체를 분석합니다.')
    weeks=(calendar.monthrange(month.year,month.month)[1]+6)//7
    default_week=min((today.day-1)//7+1,weeks) if (month.year,month.month)==(today.year,today.month) else 1
    ordinals=['첫째','둘째','셋째','넷째','다섯째']
    week=b.selectbox('분석할 주차',range(1,weeks+1),index=default_week-1,format_func=lambda n:f'{ordinals[n-1]} 주',key='trend_week')
    with st.expander('검색 주제 설정'):
        topic_text=st.text_area('검색 주제 (한 줄에 하나씩)',value='금융\n테크\n라이프스타일',key='auto_trend_topics',help='주제마다 주간과 월간 검색이 각각 한 번씩 실행됩니다.')
    seeds=list(dict.fromkeys(x.strip() for x in topic_text.splitlines() if x.strip()))
    refresh=st.button('트렌드 다시 불러오기',key='reload_auto_trend')
    if not key:
        st.info('왼쪽에서 YouTube API 키를 입력하고 Enter를 누르면 자동으로 분석합니다.')
        return
    if not 1<=len(seeds)<=10:
        st.error('검색 주제를 1개 이상 10개 이하로 입력해 주세요.')
        return
    weekly,monthly=periods(month,week)
    if weekly[1]>today or monthly[1]>today:
        st.caption('현재 날짜 이후는 아직 발생하지 않았으므로 해당 기간 결과는 진행 중인 표본입니다.')
    if 'auto_trend_cache' not in st.session_state:st.session_state.auto_trend_cache={}
    for label,(start,end),cache_name in [('주간',weekly,'week'),('월간',monthly,'month')]:
        signature=hashlib.sha256((key+'\0'+'\n'.join(seeds)+str(start)+str(end)).encode()).hexdigest()
        cache=st.session_state.auto_trend_cache.get(cache_name)
        if refresh or not cache or cache['signature']!=signature:
            try:
                with st.spinner(f'{label} 공개 검색 결과를 불러오는 중입니다...'):
                    ids=[]
                    for topic in seeds: ids.extend(search(key,topic,start,end,20,order='relevance'))
                    rows=unique_videos(fetch_videos(key,ids))
                cache={'signature':signature,'rows':rows,'at':time_now(),'error':None}
            except ValueError as exc:
                cache={'signature':signature,'rows':[],'at':time_now(),'error':str(exc)}
            st.session_state.auto_trend_cache[cache_name]=cache
        title=f'{month.year}년 {month.month}월 {ordinals[week-1]} 주' if label=='주간' else f'{month.year}년 {month.month}월'
        st.markdown(f'### {title} {label} 키워드')
        st.caption(f'{start:%Y-%m-%d} ~ {end:%Y-%m-%d} | 주제 {", ".join(seeds)} | 조회 {cache["at"]} | 중복 제거 {len(cache["rows"])}건')
        if cache['error']:
            st.error(cache['error']);continue
        stats=topic_stats(cache['rows'])
        if not stats:
            st.info('해당 기간 표본에서 두 영상 이상에 반복되는 키워드가 없습니다. 주제를 변경하거나 다시 불러와 주세요.')
            continue
        top=stats[:5]
        st.markdown('**주요 키워드:** '+ ' · '.join(f"{x['키워드']} ({x['등장 영상 수']}개 영상)" for x in top))
        st.caption('영상 ID가 같은 영상은 한 번만 집계하고 같은 단어는 한 번만 표시합니다. 반복 횟수는 표본 내 등장 영상 수입니다. 제목과 설명을 단순 분리하므로 의미를 직접 확인해 주세요.')
        with st.expander(f'{label} 키워드 근거와 다운로드',expanded=True):
            keyword_display(cache['rows'],f'{title} {label} 키워드',f'YouTube 공개 검색 / 주제 {", ".join(seeds)} / {cache["at"]}',f'{start} ~ {end}',f'auto_trend_{cache_name}')
            video_report(cache['rows'],f'{title} 검색 영상',f'{start} ~ {end}',cache['at'],f'auto_trend_videos_{cache_name}')
