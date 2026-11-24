import collections,csv,hashlib,json,os,re,shutil,zipfile
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.table import Table,TableStyleInfo
from docx import Document
from docx.shared import Inches,Pt,RGBColor,Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from policy_data import CATEGORIES,CORE,SUPPORT,rules,EXAMPLES_TSV
ROOT=Path(__file__).resolve().parent.parent; WORK=ROOT/'work'; OUT=ROOT/'outputs'
DATE=os.environ.get('SCRAPE_DATE', datetime.now().date().isoformat()); VERSION=os.environ.get('RULE_VERSION', 'v1.0')
CORPUS=json.loads((WORK/'corpus.json').read_text(encoding='utf-8'))
SUP=json.loads((WORK/'supplement.json').read_text(encoding='utf-8'))
COVER=json.loads((WORK/'coverage.json').read_text(encoding='utf-8'))
RULES=rules(); CAT={x[0]:x for x in CATEGORIES}
for r in CORPUS+SUP:
    r['title']=re.sub('<[^>]+>','',r['title']).replace('\u200b','').strip()
    r['text']=re.sub(r'\n关闭\s*$','',r['text']).strip();r['chars']=len(r['text'])
for r in SUP:
    if r['id']=='G07':r['source_kind']='部门规章'
    if r['id']=='G09':r['effective']='2026-09-30'
SOURCES={r['id']:r for r in CORPUS+SUP}
for z in RULES:
    assert all(SOURCES[s]['status']=='ok' for s in z['sources'])
    assert z['category'] in CAT
USE_IDS={s for z in RULES for s in z['sources']}
OUT.mkdir(exist_ok=True)

def screening(r):
    sid=r['id'];title=r['title']
    if r['status']!='ok':return '待核验','以图片或专题为主，正文未完整结构化，不作为独立规则证据'
    if sid in CORE:return '治理材料','与内容治理、账号管理、平台责任或执法直接相关；纳入治理材料台账'
    if sid in SUPPORT:return '背景与解读','政策解读、保护工具、行业背景或历史管理文件；与通报分开列示'
    if sid in {'S0008','S0063','S0071','S0102','S0591'}:return '相邻领域排除','主要是网络数据安全、个人信息收集或产品安全，非网络不良内容治理通报'
    return '非治理材料排除','主要为宣传活动、作品征集评选、行业动态等，与本次内容审核规则制定无直接对应'

def kind(r):
    t=r['title']
    if r['id'].startswith('G'):return r['source_kind']
    if '答记者问' in t or '解读' in t:return '政策解读/答问'
    if '规定' in t and any(x in t for x in ['发布','公布','出台']):return '规范发布/历史制度'
    if any(x in t for x in ['典型案例','曝光','查处','处罚','关停','关闭','约谈','被查处']):return '案例曝光/执法处置'
    if any(x in t for x in ['专项行动','专项整治','集中整治','部署','专项治理','联合治理']):return '专项部署/治理通知'
    if any(x in t for x in ['成效','取得','推进','开展','加强','整治','规范管理']):return '治理进展/工作成效'
    return '其他治理与背景材料'

def tags(r):
    hits=[];t=r['title'];body=r['text']
    for c,n,d,pattern,scene in CATEGORIES:
        th=len(re.findall(pattern,t,re.I));bh=len(re.findall(pattern,body,re.I));score=th*4+min(bh,8)
        if c=='C18':
            if not th and bh<5:continue
        if score>=3:hits.append((score,c,'标题+正文线索' if th else '正文关键词线索'))
    if not hits and screening(r)[0]=='治理材料':hits=[(1,'C18','综合治理兜底，需细分复核')]
    return sorted(hits,key=lambda x:(-x[0],x[1]))

for r in CORPUS:
    r['screening'],r['screen_reason']=screening(r)
    r['material_kind']=kind(r)
    r['governance_tags']=[x[1] for x in tags(r)] if r['screening'] in {'治理材料','背景与解读'} else []
    r['tag_status']='标题与正文关键词辅助初标；用于检索，需人工复核' if r['governance_tags'] else '未标'
    r['body_status']='正文已抽取（未逐字人工校对）' if r['status']=='ok' else '图片/专题待核验'
    r['content_sha256']=hashlib.sha256(r['text'].encode()).hexdigest()
COUNT=collections.Counter(r['screening'] for r in CORPUS)
CORE_ROWS=[r for r in CORPUS if r['screening']=='治理材料']
# Evidence locators identify a normalized text paragraph; they are not automatic legal conclusions.
G_ANCHORS={
('G01','C01'):['第六条'],('G01','C05'):['第六条','第七条'],('G01','C06'):['第六条','第七条'],('G01','C16'):['第七条'],('G01','C17'):['第十一条'],('G01','C18'):['第九条','第十条','第十六条'],
('G02','C05'):['第三条','第六条'],('G02','C06'):['第三条'],('G02','C07'):['第六条','第七条','第八条','第九条'],('G02','C17'):['第九条'],
('G03','C03'):['第四条','第五条','第六条','第十条'],('G03','C18'):['第九条'],('G04','C04'):['第十条','第十五条','第十八条','第三十二条','第三十三条'],('G04','C18'):['第十五条','第二十五条','第二十六条'],
('G05','C17'):['第八条','第十六条','第十七条','第二十一条'],('G05','C18'):['第二十二条'],('G06','C03'):['第八条','第十三条','第十四条','第十八条','第十九条'],('G06','C06'):['第八条','第十三条'],('G07','C18'):['第六条','第七条','第十三条','第十四条','第十六条','第二十条'],('G08','C01'):['第七条','第十条','第十三条','第十四条','第二十九条'],('G09','C09'):['第五条','第六条','第十六条','第十八条','第二十五条']}

def locator(z,r):
    lines=[p.strip() for p in r['text'].splitlines() if p.strip()]
    anchors=G_ANCHORS.get((r['id'],z['category']),[])
    if r['id']=='G03' and z['id']=='C03-04':anchors=['第十条']
    if r['id']=='G01' and z['id']=='C18-03':anchors=['第十六条']
    rule_anchors={
        ('G02','C05-03'):['第三条','第五条'],('G02','C06-01'):['第三条'],('G02','C06-02'):['第三条','第四条'],
        ('G02','C07-01'):['第二条','第七条','第八条','第九条'],('G02','C07-02'):['第五条'],('G02','C07-05'):['第六条'],
        ('G06','C07-06'):['第十四条','第十七条'],('G01','C18-02'):['第十条'],('G01','C18-04'):['第九条','第十七条'],
        ('G04','C18-02'):['第十五条','第二十五条'],('G04','C18-03'):['第二十三条','第二十六条','第三十三条']}
    anchors=rule_anchors.get((r['id'],z['id']),anchors)
    if anchors:
        matches=[(i,p) for i,p in enumerate(lines,1) if any(p.startswith(a) for a in anchors)]
        if matches:return '；'.join(f'p{i} '+re.match(r'^第[一二三四五六七八九十百]+条',p).group() for i,p in matches),matches[0][1][:24]
    scored=[]
    for i,p in enumerate(lines,1):
        if len(p)<12:continue
        score=sum((2 if len(k)>2 else 1) for k in z['keywords'] if k in p)
        if any(k in p for k in ['以下','一是','二是','重点','1.','一、']):score+=.5
        if any(k in p for k in ['有关负责人强调','各地网信部门要','请认真贯彻']):score-=2
        scored.append((score,-i,i,p))
    scored.sort(reverse=True)
    if not scored:return '全文核验',''
    best=scored[0];return f'p{best[2]}（辅助定位，需结合全文）',best[3][:24]

EVIDENCE=[]
for z in RULES:
    for sid in z['sources']:
        r=SOURCES[sid];loc,peek=locator(z,r)
        EVIDENCE.append([z['id'],z['title'],sid,r['date'],r['title'],r['url'],r.get('source_kind',kind(r)),loc,peek+'…','原文定位参考；判定条件为规则转写，边界/处置/优先级为平台建议'])

EXAMPLES=[]
for i,line in enumerate(EXAMPLES_TSV.strip().splitlines(),1):
    c,bad,handle,good,allow=line.split('|');EXAMPLES.append([f'E{i:03d}',c,CAT[c][1],bad,handle,good,allow,'构造培训样例，非真实通报事实；不可仅靠字面匹配'])

WORKFLOW=[
['P0','明确违法且重大现实危害；涉未成年人性剥削、紧迫生命危险等','即时控传播/保护性隔离→专岗复核→依法报告与必要救援','不得用临时隔离替代最终定性；保护受害人，不扩大敏感素材','平台建议优先级；无统一法定分钟数，按业务能力明确内部SLA'],
['P1','较高危害、证据明确的欺诈侵权及组织性违规','优先人工复核→内容处置→账号/营利/机构分级处置','永久关闭需充分证据及复核；账号关联不能只看设备或IP','建议设置优先队列、应急轮值和复核负责人'],
['P2','内容标注、年龄适配、低俗不良或仍待事实核验','补标纠正/限制重点推荐/适龄控制→复核后决定是否恢复或升级','证据不足不作最终违法结论；正常表达反例与语境同步审','限制推荐属于处置，也应说明理由和提供申诉'],
['P3','治理制度、审核流程、算法机制与MCN责任','责任部门整改→验收→回归抽检→规则更新','P3是治理任务类别，不代表违法行为危害一定较低','与产品、法务、安全、申诉及机构管理协同'],
['申诉','用户对内容或账号处置不服','通知理由和具体条款→便捷申诉→独立复核→结论及必要恢复','不得要求撤回投诉换恢复；保护对方隐私和依法保密资料','时限和材料清单以公开平台规则及适用规范确定'],
['留证','全流程可审计','原始内容哈希、截图录像、时间、适用条款、审核复核、处置与恢复记录','最小必要采集、权限隔离、按依据保留和删除','G03第九条不少于六个月仅针对该条情形，不扩展成所有审核日志期限'],
['质检','每次规则或模型更新','按风险和场景抽样→复核漏放与误判→申诉回归→修订培训','分别观察群体与场景误伤；抽样不能仅选择机器高分样本','建议指标：漏放率、误判率、申诉改判率、复发率、补标准确率、未成年人受限内容触达率']]

README_LINES=[
'交付说明｜网信办治理分类与互联网平台内容审核规则',
f'研究截止：{DATE}（北京时间）；版本：{VERSION}。',
'范围：用户指定“互联网内容管理”和“清朗”系列专项行动两个栏目；未设置起始年份，抓取当前官方接口可访问的全部记录。',
f'互联网内容管理：页面显示400条，接口totalRec=643，33页，643条栏目记录、642个不同URL；清朗：61条、4页、61个不同URL。两个栏目合计704条、URL去重703篇。',
f'时间范围：2013-05-02至2026-11-06。37页全部取得，703篇页面均已下载，686篇正文结构化抽取，17篇图片/专题待核验。',
f'筛选：治理材料{COUNT["治理材料"]}篇；背景与解读{COUNT["背景与解读"]}篇；其余{sum(v for k,v in COUNT.items() if k not in {"治理材料","背景与解读"})}篇为排除或待核验（详见表内原因）。',
'按URL去重；同题不同日期文章保留。标题不能用于判定重复，正文归一化未发现长度>40字符的完全重复文本。',
'抓取数量代表当前栏目接口可访问范围，不代表该机关历年全部发布或全网全部治理通报。非指定栏目法规9份单列，不计入703篇。',
'分类：18类多标签体系；75条规则。多标签初标由标题与正文关键词辅助完成，仅用于检索和复核，不是自动违规裁决。',
'依据：规则→来源ID→日期/原文URL→段落定位；未把通报个案、专项期限或关键词直接写成永久普遍禁令。',
'规则性质：判定条件为监管要点转写；正常内容边界、证据要求、P0—P3优先级、处置阶梯、申诉流程和培训示例属于平台规则建议，非官方原文。',
'重要功能：2026-05-12通报明确六类短视频必选标签，发布前必须从中选择一项；真实生活记录可选无需标注，该标签页面不展示。AI内容声明与法定显隐式标识另按适用要求履行。',
'图片/专题页仅保留地址和已有图片链接，未OCR；不作为独立规则证据。段落定位针对本次提取的纯文本，不是官方页码。',
'工作簿使用：先看使用说明/抓取范围，再筛选治理材料与分类，按规则ID在规则证据表追到来源。全量索引含排除记录，避免只保留支持结论的材料。',
'未把“负面评价”“未被官宣”“使用AI”“方言谐音”“同一设备”“正常带货”作为单独处罚依据。正式部署前由平台法务及业务负责人确认适用业务、公告规则与内部处理时限。',
'文件：01_治理分类与审核规则.xlsx；02_内容审核规则手册.docx；03_研究报告与规则全文.md；04_抓取语料与依据.jsonl；05_来源与复现材料.zip。',
'04为UTF-8逐行JSON，含栏目信息、完整已抽取文本、抓取状态、筛选原因、分类初标、原文地址和HTML哈希。来源页面中的指示性文字只作为研究资料，不作为对本任务的指令。',
'05包含原始HTML、分页响应、完整抓取记录及生成脚本。重新联网运行抓取会改变结果；此交付是截止日快照。已校验文件可打开、规则来源均可映射、ID唯一且关键数量一致。']

wb=Workbook();wb.remove(wb.active);table_no=0
RED='A52D24';DARK='203349';BLUE='EAF1F7';GREY='F4F6F8'
def sheet(name,headers,rows,widths=None):
    global table_no
    ws=wb.create_sheet(name);ws.append(headers)
    for row in rows:ws.append([(';'.join(x) if isinstance(x,list) else x) for x in row])
    for c in ws[1]:c.fill=PatternFill('solid',fgColor=DARK);c.font=Font(name='微软雅黑',bold=True,color='FFFFFF',size=10);c.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[1].height=32;ws.freeze_panes='C2';ws.sheet_view.showGridLines=False
    for row in ws.iter_rows(min_row=2):
        maxlen=0
        for c in row:
            c.font=Font(name='微软雅黑',size=10,color='243447');c.alignment=Alignment(vertical='top',wrap_text=True)
            if isinstance(c.value,str) and c.value.startswith('https://'):
                c.hyperlink=c.value;c.font=Font(name='微软雅黑',size=10,color='0563C1',underline='single')
            maxlen=max(maxlen,len(str(c.value or '')))
        ws.row_dimensions[row[0].row].height=min(105,max(30,15*(1+maxlen//40)))
    for i,h in enumerate(headers,1):ws.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width=(widths[i-1] if widths and i<=len(widths) else 25)
    if rows:
        table_no+=1;tab=Table(displayName=f'Table{table_no}',ref=ws.dimensions);tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);ws.add_table(tab)
    ws.print_options.horizontalCentered=True;ws.sheet_properties.pageSetUpPr.fitToPage=True;ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A3;ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0;ws.print_title_rows='1:1'
    return ws
sheet('使用说明',['序号','说明'],[[i,t] for i,t in enumerate(README_LINES,1)],[9,120])
coverrows=[]
for c in COVER:
    dates=[r['date'] for r in CORPUS if c['channel'] in r['channels']]
    coverrows.append([c['channel'],c['url'],c['declared_total'],c['api_total'],c['pages_expected'],c['pages_ok'],c['occurrences'],c['unique_urls'],min(dates),max(dates),'官方静态第一页+官方JsonList POST接口；每页20条'])
sheet('抓取范围',['栏目','入口URL','页面显示条数','接口totalRec','预期页数','成功页数','栏目记录数','栏目不同URL数','最早日期','最新日期','抓取方式'],coverrows,[22,62,14,14,12,12,14,16,14,14,65])
indexrows=[[r['id'],r['date'],'；'.join(r['channels']),r['title'],r['url'],r['screening'],r['material_kind'],r['body_status'],r['chars'],'；'.join(r['governance_tags']),'是' if r['id'] in USE_IDS else '否',r['screen_reason'],'；'.join(r['list_pages'])] for r in CORPUS]
sheet('全量索引',['来源ID','发布日期','所属栏目','标题','原文URL','筛选结果','材料类型','正文状态','正文字符数','治理标签初标','规则已引用','筛选说明','出现分页'],indexrows,[12,14,24,66,64,22,24,27,14,32,14,72,25])
sheet('治理材料',['来源ID','日期','标题','材料类型','治理标签初标','原文URL','规则引用','取证建议'],[[r['id'],r['date'],r['title'],r['material_kind'],'；'.join(r['governance_tags']),r['url'],'；'.join(z['id'] for z in RULES if r['id'] in z['sources']),'标签为辅助初标；规则结论须结合原文及证据条件'] for r in CORE_ROWS],[12,14,65,25,35,64,30,65])
catrows=[]
for c,n,d,pattern,scene in CATEGORIES:
    tagged=[r for r in CORE_ROWS if c in r['governance_tags']];ruleids=[z['id'] for z in RULES if z['category']==c]
    catrows.append([c,n,d,scene,len(tagged),len(ruleids),'；'.join(ruleids),'多标签初标频次，不是违规事件发生数；不同类别不能相加'])
sheet('治理分类',['类别ID','治理分类','定义与边界','重点场景','治理材料命中篇数','规则数','对应规则ID','计数解释'],catrows,[12,32,80,55,20,14,65,65])
sheet('审核规则',['规则ID','类别ID','类别','规则名称','优先级建议','判定条件（监管转写）','正常内容边界（平台建议）','处置建议','必要证据','依据来源ID','性质说明'],[[z['id'],z['category'],CAT[z['category']][1],z['title'],z['priority'],z['decision'],z['boundary'],z['action'],z['evidence'],'；'.join(z['sources']),z['nature']+'；边界、证据、优先级与处置为平台建议'] for z in RULES],[14,12,30,38,14,90,90,85,85,38,68])
sheet('规则证据',['规则ID','规则名称','来源ID','发布日期','来源标题','原文URL','来源性质','原文定位','短定位摘录','解读边界'],EVIDENCE,[14,38,12,14,68,65,25,65,42,85])
maprows=[]
for r in CORPUS:
    if r['governance_tags']:
        for score,c,reason in tags(r):maprows.append([r['id'],r['date'],r['title'],r['screening'],c,CAT[c][1],score,reason,'机器辅助初标，非违规结论',r['url']])
sheet('标签映射',['来源ID','日期','标题','筛选结果','分类ID','分类名称','检索分数','线索位置','标注状态','原文URL'],maprows,[12,14,66,22,12,32,14,25,40,65])
sheet('审核样例',['样例ID','分类ID','分类名称','触发情境（构造）','审核要点','正常边界情境（构造）','边界结论','说明'],EXAMPLES,[12,12,32,85,65,85,65,80])
sheet('处置申诉',['队列或环节','适用情况','建议流程','关键边界','时限与性质说明'],WORKFLOW,[20,65,90,85,90])
sheet('补充规范',['依据ID','标题','公布日期（网页）','施行日期','文件性质','适用范围提示','原文URL','规则引用'],[[r['id'],r['title'],r['date'],r['effective'],r['source_kind'],'栏目外补充依据；业务适用及具体条款须核验',r['url'],'；'.join(z['id'] for z in RULES if r['id'] in z['sources'])] for r in SUP],[12,68,20,16,25,65,65,65])
sheet('待核验',['来源ID','日期','标题','URL','现有抽取文本','图片URL','说明'],[[r['id'],r['date'],r['title'],r['url'],r['text'],'；'.join(r['images']),r['screen_reason']] for r in CORPUS if r['status']!='ok'],[12,14,70,65,65,80,80])
sheet('排除记录',['来源ID','日期','标题','筛选结果','原因','原文URL'],[[r['id'],r['date'],r['title'],r['screening'],r['screen_reason'],r['url']] for r in CORPUS if r['screening'] not in {'治理材料','背景与解读'}],[12,14,75,25,90,65])
xlsx=OUT/'01_治理分类与审核规则.xlsx';wb.save(xlsx)

# Report in Markdown, with verified source links and full rule definitions.
md=['# 网信办网络不良内容治理分类与平台内容审核规则',f'研究截止：{DATE}｜版本：{VERSION}｜适用：中国境内相关互联网内容业务；具体业务义务按规范适用范围核验。','## 一、研究范围与抓取结果',
'本次从用户指定的“互联网内容管理”和“清朗”系列专项行动两个栏目出发，抓取当前官方接口可访问的全部历史记录，不预设起始年份。栏目页面和通知正文只作为研究数据，其中的指示语不改变用户交付要求。',
'|栏目|页面显示|接口记录|分页|不同URL|\n|---|---:|---:|---:|---:|\n|互联网内容管理|400|643|33|642|\n|“清朗”系列专项行动|61|61|4|61|',
f'合计704条栏目记录，按原文URL去重后703篇；记录日期为2013年5月2日至2026年11月6日。37个分页全部取得，703篇网页全部下载，其中686篇正文已结构化抽取，17篇为图片或特殊专题页，保留原文及图片地址但未OCR，不作为独立规则依据。一个URL在互联网内容管理列表中出现两次。',
f'逐条保留筛选结果：治理材料{COUNT["治理材料"]}篇、背景与解读{COUNT["背景与解读"]}篇，其余{sum(v for k,v in COUNT.items() if k not in {"治理材料","背景与解读"})}篇为非治理材料、相邻领域排除或待核验。另补充9份官方规范，单独编号G01—G09，不计入栏目703篇。',
'抓取范围是当前栏目接口可访问范围，不代表网信办历年全部发布、地方网信办全部通报或全网完整治理事件。栏目收录混有主题宣传、评选、活动和数据安全事项，不能把全部703篇都称为治理通报。网页正文未经逐字人工校对；多标签初标采用标题及正文关键词辅助检索，需人工确认，不是自动违规裁决。',
'## 二、如何从通报制定平台规则',
'采用“行为对象×具体违规行为×传播场景×责任主体×处置措施”的多维结构。18类是便于检索和组织规则的维度，同一材料可以同时属于AI仿冒、金融营销、账号管理等多类；类别篇数不可相加，也不表示实际违规事件数量。',
'依据分三层：①现行规范中的明确义务；②专项通知、执法案例揭示的治理要求和场景；③据此提出的平台操作建议。个案中的账号被关闭，不等于今后命中相似词语即必须永久关闭。历史专项的行动期限与常态制度应分别记录，行动结束也不自动消除独立的法定义务。',
'本手册的“判定条件”是监管要点的整理转写；“正常内容边界、证据要求、P0—P3优先级、处置阶梯、质检和培训样例”为平台建议。对事实、权限和法律适用有争议的事项，应交专业复核；既不以机器分数替代证据，也不把未被官宣等同虚假。',
'## 三、治理分类',
'|类别|名称|治理对象和边界|主要场景|\n|---|---|---|---|']
for c,n,d,p,scene in CATEGORIES:md.append(f'|{c}|{n}|{d}|{scene}|')
md+=['## 四、优先落实的审核与产品要求',
'1. **短视频标注。**2026年5月12日的部署明确六类必选标签：“含有虚构演绎内容”“含有AI生成内容”“含有营销信息”“内容为转载”“内容为个人观点”“无需标注”。发布者必须从中选择一项方能发布；真实生活记录可选“无需标注”，该标签不在页面呈现。平台审核新增标注并分批回溯存量。短视频标签不能替代AI法定显隐式标识、专业资质或其他并行义务。[S0024]('+SOURCES['S0024']['url']+')。',
'2. **AI仿冒与虚假内容。**将真人形象声音授权、事实核验、生成合成标识分别审核。AI标签不能使诈骗、侵权或未成年人性化内容获得发布许可；AI检测概率不能单独证明内容虚假。[S0002]('+SOURCES['S0002']['url']+')、[G03]('+SOURCES['G03']['url']+')。',
'3. **未成年人分发保护。**除删除违法内容，还需识别合法但不适龄的信息，落实提示、重点环节呈现限制及不向未成年人推荐等适用要求，检查模式、学习入口、私信、群组和消费链。[G02]('+SOURCES['G02']['url']+')、[S0015]('+SOURCES['S0015']['url']+')。',
'4. **仿冒身份与高风险营销。**核验专业资质、在职状态和账号展示，重点覆盖假医生、虚假荐股、征信洗白、代理退保骗局、招聘考试包过、虚假助农及AI代言。合法债务协商、消费投诉和个人经历分享需与欺诈行为区分。[S0020]('+SOURCES['S0020']['url']+')、[S0003]('+SOURCES['S0003']['url']+')、[S0004]('+SOURCES['S0004']['url']+')。',
'5. **直播打赏、推荐和机构责任。**审核团播玩法、虚假人设、未成年人打赏、打赏排名及防禁言特权；同时覆盖热搜、搜索、弹窗、榜单和MCN策划分发责任。只删视频不足以修复不当产品机制。[S0055]('+SOURCES['S0055']['url']+')、[S0111]('+SOURCES['S0111']['url']+')、[G07]('+SOURCES['G07']['url']+')。',
'## 五、完整审核规则（75条）',
'每条均需结合具体内容、上下文、证据和适用业务认定。P0—P3为平台工作队列建议，不是法律风险等级。所有正常内容边界是操作建议，不等于对任何个案作法律豁免。']
for c,n,d,p,scene in CATEGORIES:
    md.append(f'### {c} {n}')
    for z in RULES:
        if z['category']!=c:continue
        refs='；'.join(f'[{s}]({SOURCES[s]["url"]})' for s in z['sources'])
        md.extend([f'#### {z["id"]} {z["title"]}（{z["priority"]}）',f'- 判定条件：{z["decision"]}。',f'- 正常内容边界：{z["boundary"]}。',f'- 处置建议：{z["action"]}。',f'- 必要证据：{z["evidence"]}。',f'- 依据：{refs}。性质：{z["nature"]}；边界、证据、优先级和处置为平台建议。'])
md+=['## 六、审核流程与申诉',
'建议顺序：模型与举报发现线索 → 获取完整语境和证据 → 对照规则与正常边界 → 确认适用义务 → 必要时采取比例适当的临时措施 → 人工/专业复核 → 内容和账号分级处置 → 告知理由与申诉 → 独立复核和恢复 → 质量回归。',
'账号处置应综合证据、危害、是否故意、既往有效处罚和整改表现；永久关闭、跨账号关联或机构性处罚需强化复核。共享设备、IP相同和文字雷同不足以单独证明共同控制。',
'对未核实但具有重大传播风险的信息，可设置有说明、有复核和有期限管理的临时分发措施；不能将待核验直接记录为最终谣言结论，也不应无限期搁置。真实求助、合法监督和受害者自述要有快速人工通道。',
'|环节|适用|流程建议|关键边界|\n|---|---|---|---|']
for a,b,c,d,e in WORKFLOW:md.append(f'|{a}|{b}|{c}|{d}；{e}|')
md+=['## 七、审核样例与反例（构造）','以下是依据治理行为设计的培训情境，不是对真实账号或事件的事实断言。完整36组见工作簿。','|分类|触发情境|审核要点|正常边界|\n|---|---|---|---|']
for row in EXAMPLES:md.append(f'|{row[2]}|{row[3]}|{row[4]}|{row[5]}；{row[6]}|')
md+=['## 八、九份补充规范及其日期','规范按官网当前发布文本和目录核对；“网页公布日期”和文件落款可能不同。所列施行日期均不晚于研究截止日。','|ID|规范|网页日期|施行日期|性质|\n|---|---|---|---|---|']
for r in SUP:md.append(f'|{r["id"]}|[{r["title"]}]({r["url"]})|{r["date"]}|{r["effective"]}|{r["source_kind"]}|')
md+=['## 九、验证、限制及后续维护',
'已检查703个来源ID与URL唯一、所有37页均取得、所有75条规则的来源ID均可解析且具有正文、全部补充规范具有有效文本、工作簿与Word可重新打开、JSONL可逐行解析。规则证据表提供归一化纯文本段落p号作为定位参考，其中自动推荐的定位不替代全文核验。',
'17篇待核验中包括图片正文和特殊专题页面；未做图片OCR。历史规范发布及旧案仅作背景或风险线索，正式规则应以当前有效规范和适用业务义务为准；本次不是对所有相关法律的穷尽检索。',
'后续建议持续跟踪两个栏目与相关规范公布、修订、生效及废止情况，按实际业务建立可公示的版本记录、复核时限、申诉恢复机制和误判回归集。规则上线前应进行一次代表性样本复核，重点测试真实投诉、求助、新闻报道、医疗科普、方言创作与合法AI作品的误伤。',
'## 十、来源索引（规则已引用）','完整703篇索引及筛选原因见工作簿；下表只列75条规则直接引用的来源。','|ID|日期|标题|\n|---|---|---|']
for sid in sorted(USE_IDS):
    r=SOURCES[sid];md.append(f'|{sid}|{r["date"]}|[{r["title"]}]({r["url"]})|')
mdtext='\n\n'.join(md)+'\n';mdtext=re.sub(r'(?m)(^\|[^\n]*\|)\n\n(?=\|)',r'\1\n',mdtext); mdpath=OUT/'03_研究报告与规则全文.md';mdpath.write_text(mdtext,encoding='utf-8')

# Word uses the same reviewed content; source links are actual clickable hyperlinks.
doc=Document();sec=doc.sections[0];sec.top_margin=Cm(2.0);sec.bottom_margin=Cm(1.8);sec.left_margin=Cm(2.0);sec.right_margin=Cm(2.0)
styles=doc.styles
for name in ['Normal','Body Text','List Bullet','List Number']:
    styles[name].font.name='等线';styles[name].font.size=Pt(10.5);styles[name]._element.rPr.rFonts.set(qn('w:eastAsia'),'等线')
styles['Normal'].paragraph_format.space_after=Pt(6);styles['Normal'].paragraph_format.line_spacing=1.18
for name,size in [('Title',26),('Heading 1',19),('Heading 2',15),('Heading 3',12),('Heading 4',11)]:
    styles[name].font.name='等线';styles[name]._element.rPr.rFonts.set(qn('w:eastAsia'),'等线');styles[name].font.size=Pt(size);styles[name].font.color.rgb=RGBColor.from_string(RED if name in ['Title','Heading 1'] else DARK)
    styles[name].paragraph_format.keep_with_next=True
header=sec.header.paragraphs[0];header.text='中国网信网公开材料研究  /  平台内容治理';header.style='Caption'
foot=sec.footer.paragraphs[0];foot.alignment=2;foot.add_run('研究截止 2026.11.17  ·  ')
fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');foot._p.append(fld)

def hyperlink(par,label,url):
    rel=par.part.relate_to(url,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',is_external=True)
    h=OxmlElement('w:hyperlink');h.set(qn('r:id'),rel);run=OxmlElement('w:r');props=OxmlElement('w:rPr');col=OxmlElement('w:color');col.set(qn('w:val'),'0563C1');props.append(col);run.append(props);t=OxmlElement('w:t');t.text=label;run.append(t);h.append(run);par._p.append(h)
def add_runs(p,t):
    parts=re.split(r'(\[[^\]]+\]\(https://[^)]+\)|\*\*[^*]+\*\*)',t)
    for chunk in parts:
        if not chunk:continue
        m=re.fullmatch(r'\[([^\]]+)\]\((https://[^)]+)\)',chunk)
        if m:hyperlink(p,m.group(1),m.group(2))
        elif chunk.startswith('**') and chunk.endswith('**'):p.add_run(chunk[2:-2]).bold=True
        else:p.add_run(chunk)
lines=mdtext.splitlines();i=0
while i<len(lines):
    line=lines[i].strip()
    if not line:i+=1;continue
    if line.startswith('|'):
        block=[]
        while i<len(lines):
            l=lines[i].strip()
            if l.startswith('|'):block.append(l);i+=1
            elif not l and i+1<len(lines) and lines[i+1].strip().startswith('|'):i+=1
            else:break
        vals=[[v.strip() for v in s.strip('|').split('|')] for s in block if not re.fullmatch(r'[|\s:\-]+',s)]
        if not vals:continue
        table=doc.add_table(rows=1,cols=len(vals[0]));table.style='Light Shading Accent 1'
        for j,v in enumerate(vals[0]):add_runs(table.rows[0].cells[j].paragraphs[0],v)
        trPr=table.rows[0]._tr.get_or_add_trPr();repeat=OxmlElement('w:tblHeader');trPr.append(repeat)
        for row in vals[1:]:
            cells=table.add_row().cells
            for j,v in enumerate(row):
                if j<len(cells):add_runs(cells[j].paragraphs[0],v)
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_after=Pt(3)
                    for run in p.runs:run.font.size=Pt(8)
        continue
    m=re.match(r'^(#{1,4})\s+(.*)',line)
    if m:
        level=len(m.group(1));txt=m.group(2)
        if level==1:doc.add_paragraph(txt,'Title')
        else:doc.add_heading(txt,level-1)
    elif line.startswith('- '):add_runs(doc.add_paragraph(style='List Bullet'),line[2:])
    else:add_runs(doc.add_paragraph(),line)
    i+=1
docx=OUT/'02_内容审核规则手册.docx';doc.save(docx)

jsonl=OUT/'04_抓取语料与依据.jsonl'
with jsonl.open('w',encoding='utf-8') as f:
    for r in CORPUS+SUP:f.write(json.dumps(r|{'retrieved_date':DATE,'dataset_scope':'栏目采集' if r['id'].startswith('S') else '栏目外补充规范'},ensure_ascii=False)+'\n')
(OUT/'00_交付说明.txt').write_text('\n\n'.join(README_LINES),encoding='utf-8-sig')
# Machine-readable policies and manifest live in the reproducibility archive.
(WORK/'rules.json').write_text(json.dumps(RULES,ensure_ascii=False,indent=2),encoding='utf-8')
(WORK/'evidence.csv').write_text('',encoding='utf-8')
with (WORK/'evidence.csv').open('w',newline='',encoding='utf-8-sig') as f:
    w=csv.writer(f);w.writerow(['rule_id','rule_name','source_id','date','source_title','url','source_kind','locator','short_locator_excerpt','interpretation_note']);w.writerows(EVIDENCE)
(WORK/'snapshot_README.txt').write_text('\n\n'.join(README_LINES)+'\n\n复现环境：Python 3.11+，安装openpyxl、python-docx。保留work/source_id_registry.json以稳定来源ID。先运行work/collect.py（联网，可能更新快照），再运行work/supplement.py，然后运行work/build.py。collect.py抓取分页正文；policy_data.py保存人工整理规则；build.py生成交付。首次全量运行需先读取官方接口totalRec；本次已保存完整分页响应。',encoding='utf-8')
# Validate substantive invariants and reopen actual files.
assert len(CORPUS)==703 and len({x['id'] for x in CORPUS})==703 and len({x['url'] for x in CORPUS})==703
assert sum(x['occurrences'] for x in COVER)==704 and all(x['pages_ok']==x['pages_expected'] for x in COVER)
assert len(RULES)==75 and len({x['id'] for x in RULES})==75 and len(CATEGORIES)==18 and len(EXAMPLES)==36
assert all(x['effective']<=DATE and re.fullmatch(r'\d{4}-\d{2}-\d{2}',x['effective']) for x in SUP)
assert all(x['screening']=='治理材料' for x in CORE_ROWS)
check=load_workbook(xlsx,read_only=True,data_only=False);assert check['全量索引'].max_row==704 and check['审核规则'].max_row==76 and check['待核验'].max_row==18;check.close()
assert len(Document(docx).paragraphs)>300
assert sum(1 for line in jsonl.read_text(encoding='utf-8').splitlines() if json.loads(line))==712
stats={'research_date':DATE,'column_records':704,'unique_column_pages':703,'column_page_count':37,'article_text_ok':686,'article_text_needs_review':17,'screening':dict(COUNT),'categories':18,'rules':75,'examples':36,'rule_source_links':len(EVIDENCE),'directly_cited_sources':len(USE_IDS),'supplementary_norms':9,'checks':'passed'}
(WORK/'validation.json').write_text(json.dumps(stats,ensure_ascii=False,indent=2),encoding='utf-8')
archive=OUT/'05_来源与复现材料.zip'
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for p in sorted(WORK.rglob('*')):
        if p.is_file() and '__pycache__' not in p.parts:z.write(p,p.relative_to(ROOT))
print(json.dumps(stats,ensure_ascii=False,indent=2))
print('FILES')
for p in sorted(OUT.iterdir()):print(p.name,p.stat().st_size)
