import concurrent.futures as cf
import datetime, hashlib, html, json, math, os, re, time, urllib.request, urllib.parse
from html.parser import HTMLParser
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
WORK=ROOT/'work'; OUT=ROOT/'outputs'
HEAD={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0 Safari/537.36','Accept-Language':'zh-CN,zh;q=0.9','Referer':'https://www.cac.gov.cn/wxzw/A0937index_1.htm'}
CHANNELS=[('互联网内容管理','A093704','hlwnrgl'),('清朗系列专项行动','A093711','qinglang')]
CUTOFF=os.environ.get('SCRAPE_CUTOFF', datetime.date.today().isoformat())
try:
    datetime.date.fromisoformat(CUTOFF)
except ValueError as exc:
    raise SystemExit(f'SCRAPE_CUTOFF must be YYYY-MM-DD, got {CUTOFF!r}') from exc
class Text(HTMLParser):
    def __init__(self): super().__init__(); self.parts=[]; self.skip=0; self.images=[]
    def handle_starttag(self,tag,attrs):
        a=dict(attrs)
        if tag in ('script','style'): self.skip+=1
        if tag in ('p','div','br','h1','h2','h3','li','tr'): self.parts.append('\n')
        if tag=='img': self.images.append(a.get('src','')); self.parts.append('[图片：'+a.get('alt','')+']')
    def handle_endtag(self,tag):
        if tag in ('script','style'): self.skip=max(0,self.skip-1)
        if tag in ('p','div','li','tr'): self.parts.append('\n')
    def handle_data(self,data):
        if not self.skip:self.parts.append(data)
    def text(self):return '\n'.join(x.strip() for x in ''.join(self.parts).splitlines() if x.strip())
def text(s):
    p=Text();p.feed(s);return p.text()
def fetch(url,data=None):
    last=None
    for k in range(3):
        try:
            req=urllib.request.Request(url,data=data,headers=HEAD|({'X-Requested-With':'XMLHttpRequest'} if data else {}))
            with urllib.request.urlopen(req,timeout=32) as r:
                b=r.read(); charset=r.headers.get_content_charset() or 'utf-8'
                return b.decode(charset,errors='replace')
        except Exception as e:last=e;time.sleep(.3*(k+1))
    raise last
def parse_list(s):
    m=re.search(r"id=['\"]loadingInfoPage['\"][^>]*>(.*?)<div[^>]+id=['\"]Page",s,re.S|re.I)
    area=m.group(1) if m else s
    found=[]
    for u,title,date in re.findall(r'<li><h5><a href=(\S+) target=_blank title=[\"](.*?)[\"]>.*?</a></h5><div class=[\"]times[\"]>(.*?)</div></li>',area,re.S):
        found.append({'url':urllib.parse.urljoin('https://www.cac.gov.cn/',u.strip('\"\'')),'title':html.unescape(title),'date':date})
    return found
def page_job(code,slug,n):
    url=f'https://www.cac.gov.cn/wxzw/{slug}/{code}index_{n}.htm'
    try:
        s=fetch(url); items=parse_list(s)
        if items:return {'page':n,'method':'static','items':items,'url':url,'html':s}
    except Exception as e: err=str(e)
    form=urllib.parse.urlencode(dict(channelCode=code,perPage=20,pageno=n,condition='0',fuhao='=',value='')).encode()
    for api,data in [('https://www.cac.gov.cn/cms/JsonList',form),('https://www.cac.gov.cn/cms/JsonList?'+form.decode(),None)]:
        try:
            s=fetch(api,data); j=json.loads(s);items=[{'url':urllib.parse.urljoin('https://www.cac.gov.cn/',x['infourl']),'title':html.unescape(x['topic']),'date':x['pubtime'][:10]} for x in j['list']]
            if items:return {'page':n,'method':'json','items':items,'url':api,'html':s}
        except Exception as e:err=str(e)
    return {'page':n,'items':[],'error':err,'url':url}
def article(item):
    path=WORK/'html'/f"{hashlib.sha256(item['url'].encode()).hexdigest()[:16]}.html"
    try:
        s=path.read_text(encoding='utf-8') if path.exists() else fetch(item['url'])
        path.write_text(s,encoding='utf-8')
        m=re.search(r'<(?:div|DIV)\s+id\s*=\s*[\"\']?BodyLabel[\"\']?\s*[^>]*>(.*?)<!--\s*footer',s,re.S)
        method='BodyLabel'
        if not m:
            m=re.search(r'<!--\s*(?:正文|文章正文|content)[^>]*-->(.*?)<!--',s,re.S|re.I);method='comment'
        area=m.group(1) if m else ''
        if area: area=re.split(r'<div[^>]+class=[\"\'](?:share|bdshare|editor|article-bottom)',area,flags=re.I)[0]
        p=Text();p.feed(area);body=p.text()
        imgs=[urllib.parse.urljoin(item['url'],u) for u in p.images if u]
        return item|{'status':'ok' if len(body)>30 else 'needs_review','extract_method':method if m else 'not_found','text':body,'images':imgs,'sha256_html':hashlib.sha256(s.encode()).hexdigest(),'html_file':str(path.relative_to(ROOT)),'chars':len(body)}
    except Exception as e:return item|{'status':'failed','error':str(e),'text':'','images':[],'chars':0}
def main():
    (WORK/'html').mkdir(parents=True,exist_ok=True);(WORK/'lists').mkdir(parents=True,exist_ok=True)
    records={};coverage=[]
    for name,code,slug in CHANNELS:
        u=f'https://www.cac.gov.cn/wxzw/{slug}/{code}index_1.htm';s=fetch(u)
        (WORK/'lists'/f'{code}_1.html').write_text(s,encoding='utf-8')
        total=int(re.search(r"total:'(\d+)'",s).group(1))
        form=urllib.parse.urlencode(dict(channelCode=code,perPage=20,pageno=1,condition='0',fuhao='=',value='')).encode()
        probe_text=fetch('https://www.cac.gov.cn/cms/JsonList',form)
        (WORK/'lists'/f'{code}_probe.json').write_text(probe_text,encoding='utf-8')
        probe=json.loads(probe_text);api_total=int(probe['totalRec']);n=math.ceil(max(total,api_total)/20)
        pages=[{'page':1,'method':'static','items':parse_list(s),'url':u}]
        with cf.ThreadPoolExecutor(max_workers=4) as pool:
            for result in pool.map(lambda i:page_job(code,slug,i),range(2,n+1)):
                if 'html' in result:(WORK/'lists'/f'{code}_{result["page"]}.txt').write_text(result.pop('html'),encoding='utf-8')
                pages.append(result)
        occurrences=0; channel_urls=set()
        for pg in pages:
            for row in pg['items']:
                occurrences+=1;channel_urls.add(row['url'])
                if row['date']>CUTOFF:continue
                if row['url'] not in records:records[row['url']]=row|{'channels':[],'list_pages':[]}
                if name not in records[row['url']]['channels']:records[row['url']]['channels'].append(name)
                records[row['url']]['list_pages'].append(f'{code}:{pg["page"]}')
        coverage.append({'channel':name,'url':u,'declared_total':total,'api_total':api_total,'pages_expected':n,'pages_ok':sum(bool(x['items']) for x in pages),'occurrences':occurrences,'unique_urls':len(channel_urls),'pages':pages})
        print(name,total,'pages',len(pages),'unique',len(channel_urls),'errors',[(x['page'],x.get('error')) for x in pages if not x['items']],flush=True)
    (WORK/'coverage.json').write_text(json.dumps(coverage,ensure_ascii=False,indent=2),encoding='utf-8')
    rows=sorted(records.values(),key=lambda r:(r['date'],r['url']),reverse=True)
    old=json.loads((WORK/'index.json').read_text(encoding='utf-8')) if (WORK/'index.json').exists() else []
    registry_path=WORK/'source_id_registry.json'
    ids=json.loads(registry_path.read_text(encoding='utf-8')) if registry_path.exists() else {}
    ids.update({r['url']:r['id'] for r in old});next_id=max([int(v[1:]) for v in ids.values()]+[0])
    for r in rows:
        if r['url'] in ids:r['id']=ids[r['url']]
        else:next_id+=1;r['id']=f'S{next_id:04d}'
    (WORK/'index.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
    ids.update({r['url']:r['id'] for r in rows});registry_path.write_text(json.dumps(ids,ensure_ascii=False,indent=2),encoding='utf-8')
    done=[]
    with cf.ThreadPoolExecutor(max_workers=6) as pool:
        for i,r in enumerate(pool.map(article,rows),1):
            done.append(r)
            if i%25==0 or i==len(rows):
                (WORK/'corpus.json').write_text(json.dumps(done,ensure_ascii=False,indent=2),encoding='utf-8')
                print('articles',i,'/',len(rows),'ok',sum(x['status']=='ok' for x in done),'review',sum(x['status']=='needs_review' for x in done),'failed',sum(x['status']=='failed' for x in done),flush=True)
    assert len({r['url'] for r in done})==len(done)
    assert all(r['date']<=CUTOFF for r in done)
    print('DATE RANGE',min(r['date'] for r in done),max(r['date'] for r in done),flush=True)
if __name__=='__main__':main()

