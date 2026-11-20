import datetime as dt
import re
import sqlite3
import unicodedata
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import requests

SEOUL = ZoneInfo('Asia/Seoul')
BASE = 'https://www.googleapis.com/youtube/v3/'

def video_id(value):
    value = value.strip()
    if re.fullmatch(r'[A-Za-z0-9_-]{11}', value): return value
    u = urlparse(value)
    host = (u.hostname or '').lower()
    if host in ('youtu.be','www.youtu.be'): candidate = u.path.strip('/').split('/')[0]
    elif host in ('youtube.com','www.youtube.com','m.youtube.com'):
        parts = u.path.strip('/').split('/')
        candidate = parse_qs(u.query).get('v',[''])[0] if u.path == '/watch' else (parts[1] if len(parts)>1 and parts[0] in ('shorts','embed','live') else '')
    else: candidate = ''
    if re.fullmatch(r'[A-Za-z0-9_-]{11}',candidate): return candidate
    raise ValueError(f'영상 주소 또는 ID를 확인해 주세요: {value}')

def channel_filter(value):
    value=value.strip()
    if re.fullmatch(r'UC[A-Za-z0-9_-]{22}',value): return {'id':value}
    if value.startswith('@') and re.fullmatch(r'@[A-Za-z0-9_.-]{3,30}',value): return {'forHandle':value}
    u=urlparse(value)
    if (u.hostname or '').lower() not in ('youtube.com','www.youtube.com','m.youtube.com'): raise ValueError(f'채널 주소 또는 ID를 확인해 주세요: {value}')
    p=u.path.strip('/').split('/')
    if len(p)==2 and p[0]=='channel' and re.fullmatch(r'UC[A-Za-z0-9_-]{22}',p[1]): return {'id':p[1]}
    if len(p)==1 and re.fullmatch(r'@[A-Za-z0-9_.-]{3,30}',p[0]): return {'forHandle':p[0]}
    raise ValueError('채널 URL은 /channel/ID 또는 /@핸들 형식을 사용해 주세요.')

def api(key, endpoint, **params):
    key=key.strip()
    if not key:
        raise ValueError('YouTube Data API 키를 입력해 주세요.')
    try:
        response=requests.get(BASE+endpoint,params={**params,'key':key},timeout=20)
    except requests.exceptions.ProxyError:
        raise ValueError('프록시 서버에 연결할 수 없습니다. 테스트 환경에서 실행 중이라면 서버를 종료하고 폴더의 실행.bat을 직접 실행해 주세요. 회사 네트워크라면 프록시 설정을 확인해 주세요.') from None
    except requests.exceptions.SSLError:
        raise ValueError('보안 연결을 확인할 수 없습니다. PC의 인증서 또는 회사 네트워크 설정을 확인해 주세요.') from None
    except requests.exceptions.Timeout:
        raise ValueError('Google API의 응답 시간이 초과되었습니다. 네트워크를 확인한 뒤 다시 시도해 주세요.') from None
    except requests.RequestException:
        raise ValueError('Google API에 연결하지 못했습니다. 인터넷 및 네트워크 설정을 확인해 주세요.') from None
    if not response.ok:
        reasons=[]
        try:
            error=response.json().get('error',{})
            reasons.extend(str(x.get('reason','')) for x in error.get('errors',[]) if isinstance(x,dict))
            reasons.extend(str(x.get('reason','')) for x in error.get('details',[]) if isinstance(x,dict))
            reasons.append(str(error.get('status','')))
            message=str(error.get('message','')).lower()
        except (ValueError,TypeError,AttributeError): message=''
        normalized={x.lower() for x in reasons if x}
        if normalized & {'api_key_invalid','keyinvalid'} or 'api key not valid' in message:
            detail='API 키가 올바르지 않습니다. Google Cloud의 사용자 인증 정보에서 키 값을 다시 복사해 입력해 주세요. 프로젝트 ID나 클라이언트 ID는 API 키가 아닙니다.'
        elif normalized & {'api_key_service_blocked','accessnotconfigured','service_disabled','service_not_enabled'}:
            detail='이 키에서 YouTube Data API v3를 사용할 수 없습니다. 해당 프로젝트의 API 사용 설정과 키의 API 제한을 확인해 주세요.'
        elif normalized & {'api_key_http_referrer_blocked','api_key_ip_address_blocked','iprefererblocked','refererblocked'}:
            detail='API 키의 애플리케이션 제한이 현재 연결을 허용하지 않습니다. Python 앱은 웹사이트 제한 대신 실행 환경에 맞는 제한이 필요합니다.'
        elif normalized & {'quotaexceeded','dailylimitexceeded','resource_exhausted','ratelimitexceeded'}:
            detail='API 할당량 또는 요청 한도를 초과했습니다. Google Cloud 할당량과 요청 횟수를 확인해 주세요.'
        elif response.status_code==400:
            detail='Google이 요청을 유효하지 않다고 응답했습니다. 입력값과 API 키를 확인해 주세요.'
        elif response.status_code in (401,403):
            detail='API 접근이 거부되었습니다. API 사용 설정, 키 제한 및 할당량을 확인해 주세요.'
        else:
            detail='Google API 요청에 실패했습니다. 잠시 후 다시 시도해 주세요.'
        safe_reasons=[x for x in reasons if re.fullmatch(r'[A-Za-z][A-Za-z0-9_]{0,60}',x)]
        reason_text=f', 오류 코드: {safe_reasons[0]}' if safe_reasons else ''
        raise ValueError(f'YouTube API 요청 실패 (HTTP {response.status_code}{reason_text}). {detail}')
    return response.json()

def fetch_videos(key, ids):
    out=[]
    ids=list(dict.fromkeys(ids))
    for i in range(0,len(ids),50):
        for x in api(key,'videos',part='snippet,statistics,contentDetails',id=','.join(ids[i:i+50])).get('items',[]):
            s=x.get('snippet',{}); q=x.get('statistics',{}); c=x.get('contentDetails',{})
            out.append({'영상ID':x['id'],'제목':s.get('title',''), '설명':s.get('description',''), '채널':s.get('channelTitle',''), '채널ID':s.get('channelId',''), '발행일':s.get('publishedAt',''), '길이':c.get('duration',''), '썸네일':s.get('thumbnails',{}).get('medium',{}).get('url',''), '조회수':int(q['viewCount']) if 'viewCount' in q else None,'좋아요':int(q['likeCount']) if 'likeCount' in q else None,'댓글수':int(q['commentCount']) if 'commentCount' in q else None,'URL':'https://www.youtube.com/watch?v='+x['id']})
    return out

def popular_chart(key, limit=50):
    data=api(key,'videos',part='snippet,statistics,contentDetails',chart='mostPopular',regionCode='KR',maxResults=min(50,max(1,limit)))
    rows=[]
    for x in data.get('items',[]):
        s=x.get('snippet',{}); stats=x.get('statistics',{}); content=x.get('contentDetails',{})
        def num(name):
            value=stats.get(name)
            try: return int(value) if value is not None else None
            except (TypeError,ValueError): return None
        category=s.get('categoryId','')
        category_names={'1':'영화 및 애니메이션','2':'자동차','10':'음악','15':'동물','17':'스포츠','19':'여행','20':'게임','22':'인물·블로그','23':'코미디','24':'엔터테인먼트','25':'뉴스·정치','26':'노하우·스타일','27':'교육','28':'과학·기술'}
        rows.append({'영상ID':x['id'],'제목':s.get('title',''),'설명':s.get('description',''),'채널':s.get('channelTitle',''),'발행일':s.get('publishedAt',''),'분류':category_names.get(category,'기타 / 미확인'),'분류 ID':category,'길이':content.get('duration',''),'썸네일':s.get('thumbnails',{}).get('medium',{}).get('url',''),'조회수':num('viewCount'),'좋아요':num('likeCount'),'댓글수':num('commentCount'),'URL':'https://www.youtube.com/watch?v='+x['id']})
    return rows

def search(key, keyword, start, end, limit=25, order='date'):
    start_dt=dt.datetime.combine(start,dt.time.min,SEOUL).astimezone(dt.timezone.utc)
    end_dt=dt.datetime.combine(end+dt.timedelta(days=1),dt.time.min,SEOUL).astimezone(dt.timezone.utc)-dt.timedelta(microseconds=1)
    params=dict(part='snippet',type='video',publishedAfter=start_dt.isoformat(),publishedBefore=end_dt.isoformat(),maxResults=limit,order=order,regionCode='KR',relevanceLanguage='ko')
    if keyword and keyword.strip(): params['q']=keyword.strip()
    data=api(key,'search',**params)
    return [x['id']['videoId'] for x in data.get('items',[]) if 'videoId' in x.get('id',{})]

def channel_videos(key,value,limit=50):
    data=api(key,'channels',part='contentDetails',**channel_filter(value))
    if not data.get('items'): raise ValueError('채널을 찾지 못했습니다. 채널 ID 또는 핸들을 확인해 주세요.')
    playlist=data['items'][0]['contentDetails']['relatedPlaylists']['uploads']
    items=api(key,'playlistItems',part='contentDetails',playlistId=playlist,maxResults=limit).get('items',[])
    return [x['contentDetails']['videoId'] for x in items]

def comments(key,vid):
    data=api(key,'commentThreads',part='snippet',videoId=vid,maxResults=30,textFormat='plainText',order='relevance')
    return [x['snippet']['topLevelComment']['snippet'].get('textDisplay','') for x in data.get('items',[])]

def load_workbook(source):
    wb=openpyxl.load_workbook(source,data_only=True)
    if '1. 신규' not in wb.sheetnames: raise ValueError('엑셀에서 “1. 신규” 시트를 찾을 수 없습니다.')
    ws=wb['1. 신규']; records=[]; issues=[]
    headers=[ws.cell(4,i).value for i in range(1,ws.max_column+1)]
    id_col=next((i+1 for i,h in enumerate(headers) if h and str(h).strip().lower() in ('영상id','영상 id','video id','url','영상 url')),None)
    for row in range(5,ws.max_row+1):
        date=ws.cell(row,3).value
        if not isinstance(date,(dt.date,dt.datetime)): continue
        v=lambda c: ws.cell(row,c).value
        vid=''
        if id_col and v(id_col):
            try: vid=video_id(str(v(id_col)))
            except ValueError: issues.append(f'{row}행: 영상 ID 또는 URL 형식 확인 필요')
        r={'행':row,'발행일':pd.Timestamp(date),'재생목록':v(4),'제목':v(5),'소재':v(6),'조회수':v(7),'광고비':v(8),'광고조회수':v(10),'CPV':v(11),'좋아요':v(12),'댓글수':v(13),'광고유입비율':v(21),'검색유입비율':v(22),'추천유입비율':v(23),'평균조회율':v(31),'영상ID':vid}
        if not isinstance(r['소재'],str) or not r['소재'].strip():
            issues.append(f'{row}행: 소재 누락 또는 잘못된 형식 ({r["소재"]})')
            r['소재']=None if r['소재'] is None else str(r['소재'])
        for field in ('조회수','광고비','광고조회수','CPV','좋아요','댓글수','광고유입비율','검색유입비율','추천유입비율','평균조회율'):
            if r[field] not in (None,'-') and not isinstance(r[field],(int,float)): issues.append(f'{row}행: {field} 숫자 형식 확인 필요')
            if not isinstance(r[field],(int,float)): r[field]=None
        records.append(r)
    df=pd.DataFrame(records)
    if not df.empty:
        dup=df[df.duplicated(['제목','발행일'],keep=False)]
        if not dup.empty: issues.append('중복 제목·발행일 후보 행: '+', '.join(map(str,dup['행'].tolist())))
        title=str(ws['B2'].value or '')
        match=re.search(r'(\d{2})-(\d{2})-(\d{2})',title)
        if match:
            cutoff=pd.Timestamp(2000+int(match[1]),int(match[2]),int(match[3]))
            later=df[df['발행일']>cutoff]
            if not later.empty: issues.append(f'시트 기준일 {cutoff.date()} 이후 발행된 데이터 {len(later)}건이 포함되어 있습니다.')
    issues.append('요약 시트의 일부 평균은 전체 연도를 참조할 수 있으므로 분석 화면은 원본 영상 행을 직접 집계합니다.')
    return df, issues, wb.sheetnames

def candidates(df, public):
    result=[]
    for _,r in df.iterrows():
        for p in public:
            if r['영상ID'] and r['영상ID']==p['영상ID']: result.append({'엑셀 행':r['행'],'엑셀 제목':r['제목'],'공개 영상':p['제목'],'영상ID':p['영상ID'],'판정':'ID 일치'})
            elif r['제목'] and SequenceMatcher(None,str(r['제목']).lower(),p['제목'].lower()).ratio()>=0.65: result.append({'엑셀 행':r['행'],'엑셀 제목':r['제목'],'공개 영상':p['제목'],'영상ID':p['영상ID'],'판정':'제목 후보 - 확인 필요'})
    return result

def keywords(strings):
    stop={'현대카드','현대커머셜','그리고','있는','대한','하는','에서','이것','영상','합니다','것이','우리','정말','입니다','the','and','you'}
    words=[]
    for s in strings:
        words += [w.lower() for w in re.findall(r'[가-힣]{2,}|[A-Za-z]{3,}',str(s)) if w.lower() not in stop]
    return Counter(words).most_common(15)

def weekly_keyword_stats(rows):
    """Count each phrase once per distinct video, with traceable public examples."""
    stop={'그리고','이렇게','이것이','하는','에서','있는','대한','이번','영상','이것','저것','무엇','정말','합니다','입니다','the','and','you','shorts','short','with','that','this','www','https','com','현대카드','현대커머셜'}
    counts=Counter(); examples={}
    for row in rows:
        title=str(row.get('제목') or '')
        tokens=re.findall(r'[가-힣]{2,}|[A-Za-z]{3,}|[A-Za-z0-9]*[A-Za-z][A-Za-z0-9]*',title.lower())
        tokens=[w for w in tokens if len(w)>=2 and w not in stop]
        phrases=set(tokens)
        phrases.update(tokens[i]+' '+tokens[i+1] for i in range(len(tokens)-1))
        for word in phrases:
            counts[word]+=1
            examples.setdefault(word,[]).append({'제목':title,'URL':row.get('URL','')})
    return [{'키워드':word,'등장 영상 수':n,'표본 영상 수':len(rows),'예시 제목':examples[word][0]['제목'],'예시 URL':examples[word][0]['URL']} for word,n in counts.most_common() if n>=2]

def suggest_for_video(stats,video,extra=''):
    title=str(video.get('제목') or '').lower()
    desc=str(video.get('설명') or '').lower()
    text=' '.join((title,desc,str(extra).lower()))
    found=[]
    for item in stats:
        word=item['키워드'].lower()
        if word in text:
            source='제목 일치' if word in title else ('설명 일치' if word in desc else '직접 입력 주제 일치')
            found.append({**item,'관련 근거':source})
    return sorted(found,key=lambda x:(0 if x['관련 근거']=='제목 일치' else 1,-x['등장 영상 수'],x['키워드']))

def unique_videos(rows):
    """Deduplicate identifiable videos while retaining entries without identifiers."""
    seen=set(); result=[]
    for row in rows:
        identity=row.get('영상ID') or row.get('URL')
        if identity:
            if identity in seen: continue
            seen.add(identity)
        result.append(row)
    return result

def topic_stats(rows):
    """Title and description mentions, counted once per video, with a source example."""
    stop={'현대카드','현대커머셜','그리고','이렇게','이것이','하는','에서','있는','대한','이번','영상','이것','저것','무엇','정말','합니다','입니다','the','and','you','shorts','short','with','that','this','www','https','com','구독','좋아요','시청','공식','채널'}
    rows=unique_videos(rows)
    count=Counter(); samples={}
    for item in rows:
        fields={'제목':str(item.get('제목') or ''),'설명':str(item.get('설명') or '')[:1500]}
        seen={}
        for field,text in fields.items():
            text=re.sub(r'https?://\S+',' ',unicodedata.normalize('NFKC',text).casefold())
            tokens=[x for x in re.findall(r'[가-힣]{2,}|[a-z][a-z0-9]{2,}',text) if x not in stop]
            for term in sorted(set(tokens)): seen.setdefault(term,field)
        for term,field in seen.items():
            count[term]+=1
            samples.setdefault(term,(item.get('제목',''),item.get('URL',''),field))
    return [{'키워드':word,'등장 영상 수':n,'표본 영상 수':len(rows),'예시 제목':samples[word][0],'예시 URL':samples[word][1],'예시 위치':samples[word][2]} for word,n in count.most_common() if n>=2]

def save_snapshots(rows,db='snapshots.sqlite3'):
    now=dt.datetime.now(dt.timezone.utc).isoformat()
    with sqlite3.connect(db) as conn:
        conn.execute('CREATE TABLE IF NOT EXISTS snapshots (video_id TEXT, captured_at TEXT, views INTEGER, published_at TEXT, title TEXT)')
        conn.executemany('INSERT INTO snapshots VALUES (?,?,?,?,?)',[(r['영상ID'],now,r['조회수'],r['발행일'],r['제목']) for r in rows])

def history(db='snapshots.sqlite3'):
    if not Path(db).exists(): return pd.DataFrame()
    with sqlite3.connect(db) as conn: return pd.read_sql_query('SELECT * FROM snapshots ORDER BY captured_at DESC',conn)
