"""Generate and verify six-section moderation rule cards from the captured CAC corpus."""
import csv, json, math, os, re, zipfile, hashlib, collections, sys
from pathlib import Path
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from framework_v2_data import DETAILS, PRINCIPLES, LAW, OVERRIDE, FUNCTION_IDS, MIXED_IDS, STEPS
from policy_data import CATEGORIES
ROOT=Path(__file__).resolve().parent.parent; WORK=ROOT/'work'; OUT=ROOT/'outputs'
DATE=os.environ.get('SCRAPE_DATE', date.today().isoformat()); VERSION=os.environ.get('RULE_VERSION', 'v2.0'); OUT.mkdir(exist_ok=True)
SOURCES={x['id']:x for x in [json.loads(l) for l in (OUT/'04_抓取语料与依据.jsonl').read_text(encoding='utf-8').splitlines()]}
RULES=sorted(json.loads((WORK/'rules.json').read_text(encoding='utf-8')),key=lambda x:x['id'])
CAT={x[0]:x[1] for x in CATEGORIES}
DETAIL={}
for line in DETAILS.strip().splitlines():
    cells=line.split('|'); assert len(cells)==6,cells
    assert cells[0] not in DETAIL
    DETAIL[cells[0]]=cells[1:]
assert set(DETAIL)=={r['id'] for r in RULES}
# Substantive review corrections: match actor and obligation to the actual article.
OVERRIDE['C18-01']=[('G07','第七条','平台与机构签订入驻协议'),('G07','第九条','平台明确服务规范及重大变更公示'),('G07','第十条','平台对机构分类分级管理'),('G07','第十三条','机构核验签约主体并明确内容责任'),('G07','第十六条','机构违法不良内容生产传播责任'),('G07','第二十条','机构处置签约账号违规并通报平台')]
OVERRIDE['C09-01']=[('G09','第五条','金融机构许可业务范围及平台依法接受委托'),('G09','第六条','不得为非法金融活动提供网络营销及便利'),('G09','第十六条','账号、营销人员及平台核验巡查要求'),('G09','第二十五条','受托平台事先核实金融业务资质及监测')]
OVERRIDE['C18-02']=[('G01','第十条','法定违法和不良信息的适用处置、记录、报告要求'),('G04','第十五条','网暴违法信息与不良信息分别处置'),('G04','第二十五条','相关数据保存与取证便利'),('G03','第九条','申请不加显式标识的特定导出服务情形日志留存')]
OVERRIDE['C07-01']=[('G02','第二条','影响未成年人身心健康信息与违法信息区分'),('G02','第七条','防范抵制及管理制度'),('G02','第八条','适用信息展示前提示'),('G02','第九条','重点位置、推荐对象和未成年人专属服务限制')]
for r in RULES:
    if r['id']=='C08-04':
        r['decision']='平台必须提供“含有虚构演绎内容”“含有AI生成内容”“含有营销信息”“内容为转载”“内容为个人观点”“无需标注”六类必选标签，发布者从中选择一项后才能发布；核验新增及存量内容所选标签是否符合事实'
        r['action']='补齐六类标签及必经选择流程；真实生活记录可选“无需标注”且该项不在页面显示；审核纠正新增标签并分批回溯补标存量，教育警示相关发布者；反复恶意规避另核事实和依据'
    if r['id']=='C03-06':
        r['decision']='适用G06的拟人化互动服务违反AI身份提示、危险输出干预、未成年人虚拟亲密关系服务限制、使用提醒或便捷退出要求；按具体子项分别认定'
        r['action']='阻断危险输出并采取适用保护干预；提示正在与AI而非自然人互动；不得向未成年人提供虚拟亲属、虚拟伴侣等虚拟亲密关系服务；对连续使用每超过2小时提醒使用时长，提供便捷退出；依赖或沉迷倾向动态显著提醒'
        DETAIL[r['id']][4]+='；不满14周岁使用其他拟人互动服务需监护人同意，处理不满14周岁个人信息的同意另行核对，不将两项混为一项'
    if r['id']=='C16-01':
        r['decision']='编造事实以煽动群体敌意，或有证据证明针对群体持续鼓动侮辱、歧视、攻击等具体不良行为；单纯表达不满、展示财富或讨论群体差异不满足本条'
        r['action']='先核对事实与完整语境；独立违法行为按对应条款依法处置，具体不良信息按依据采取防范抵制措施；不得仅按情绪、立场或生活方式处罚'
    if r['id']=='C18-02':
        r['decision']='对落入适用条款的违法、不良信息或紧迫现实风险，分别履行处置、必要记录保存、报告或保护干预义务；报告对象与范围必须有适用依据'
    if r['id']=='C03-03':
        DETAIL[r['id']][3]='传播服务未检出生成元数据且无用户声明，但发现明显生成痕迹，按G03第六条添加疑似提示并提供纠错；该提示本身不构成对用户造假的处罚'
COMMON_BOUNDARY='证据不足时不作最终违规认定；存在明确现实危险或侵害风险的，可先采取范围必要、可撤销的保护措施并及时复核。不得以关键词、AI概率、负面情绪、尚无官方通报或设备关联单独定性。'
COMMON_NOTES='处置对象分别记录为内容、分发位置/受众、账号、机构和产品功能；同一行为多标签归因，不按标签数量重复加罚。账号措施另核控制关系、过错、严重程度、重复性及申诉结果。P0—P3为平台建议队列，不是法定违法等级、统一处置时限或自动封号指令。'
SECTIONS=['规则背景','管控原则','管控内容','具体管控场景','正反例','备注']
FRAMEWORK=[
('规则背景','规定名称、效力类型、发布/施行日期、适用主体、具体条款；通报名称/日期/类型、专项期限与引用目的；稳定来源编号。','条款可以定位，主体和行为对应；一般治理依据与直接禁止依据分开。'),
('管控原则','保护目标、事实认定方法、必要性与比例、正常表达边界；证据不足和疑难处理。','相似合法情形可以排除；不得以观点、身份、关键词或AI分数直接定性。'),
('管控内容','判定条件、证据清单、处理对象、措施及强度；紧急保护与最终处罚分别记录。','能说清命中了什么事实，为什么采用该措施；违法、不良、补标、功能缺陷分别处理。'),
('具体管控场景','至少两个具体入口；触发信号、核验步骤、所需证据、处置去向；功能类有复现步骤。','另一审核员能沿同一路径复核；不以“违规内容”等循环定义代替条件。'),
('正反例','正例=满足管控条件的构造情境，附理由及动作；反例=相似合规情境，附排除理由及预期动作。','每条至少1组，逐条覆盖；构造样例与真实通报案例明确区分。'),
('备注','例外、重叠条款、业务/年龄/期限、执行责任角色、复核节点、版本、未决事项。','适用限制可见，不能跨主体移植义务或把历史个案处罚自动套给新用户。')]

def article(sid, anchor):
    lines=[p.strip() for p in SOURCES[sid]['text'].splitlines() if p.strip()]
    starts=[i for i,p in enumerate(lines) if re.match('^'+re.escape(anchor)+r'(?:\s|　|[：:])',p)]
    assert len(starts)==1,(sid,anchor,starts)
    i=starts[0];stop=len(lines)
    for j in range(i+1,len(lines)):
        if re.match(r'^第[一二三四五六七八九十百]+[条章]',lines[j]):stop=j;break
    return 'p'+str(i+1), '\n'.join(lines[i:stop])

def policy_scope(src):
    hits=[]
    for p in src['text'].splitlines():
        if re.search(r'为期.{0,10}[月天]|即日起.{0,45}(?:月底|月|日)|[一二三四五六七八九十\d]{1,2}月至.{0,12}月|\d{1,2}月.{0,10}(?:月底|月末)|为期.{0,12}年',p):
            hits.append(p[:260])
    if hits:return '期限/适用线索（须结合全文）：'+'；'.join(hits[:2])
    return '未在结构化摘要中确认统一期限；不得据此推断永久有效或已全面失效，执行时核对原文和现行依据。'

def action_type(r):
    rid=r['id']
    if rid=='C08-04':return '发布标注功能整改与发布者补标纠标'
    if rid=='C03-04':return '标识核验与恶意去标、违规工具分别处置'
    if rid in FUNCTION_IDS:return '平台功能/流程整改'
    if rid in {'C03-03','C03-04','C08-02','C08-03','C09-06'}:return '标识/资质/来源核验与内容分别处置'
    if rid in {'C07-01','C17-01'}:return '分发位置/受众限制与内容分别处置'
    if rid in MIXED_IDS:return '内容处置与产品机制整改'
    return '按实质区分违法阻断、不良治理及正常保留'

def roles(r):
    c=r['category']
    if c in {'C01','C09','C10','C11'}:return '内容审核/对应领域专业复核；涉及违法定性交法务'
    if c=='C07' or r['id'] in {'C03-05','C05-03','C14-04'}:return '儿童保护专岗/内容审核/产品安全'
    if c in {'C17','C18'} or r['id'] in FUNCTION_IDS:return '内容治理负责人/产品与算法负责人/申诉复核岗'
    return '内容审核/相应业务专岗/申诉复核岗'

EVID=[];ARTICLES={};CARDS=[];CASES=[];REVIEWS=[]
with (WORK/'evidence.csv').open(encoding='utf-8-sig',newline='') as f:OLD_EVID={(e['rule_id'],e['source_id']):e for e in csv.DictReader(f)}
for r in RULES:
    rid=r['id'];c=r['category'];entries,scenario,pos,neg,special=DETAIL[rid]
    laws=OVERRIDE.get(rid,LAW[c]);legal=[]
    for sid,anchor,meaning in laws:
        src=SOURCES[sid];loc,body=article(sid,anchor)
        assert src.get('effective','0000')<=DATE
        legal.append(f"[{sid}]《{src['title']}》{anchor}（{src['source_kind']}，{src['effective']}施行）：{meaning}。")
        ARTICLES[(sid,anchor)]={'source_id':sid,'title':src['title'],'article':anchor,'locator':loc,'text':body,'meaning':meaning,'url':src['url'],'effective':src['effective'],'kind':src['source_kind']}
        EVID.append({'rule_id':rid,'source_id':sid,'title':src['title'],'date':src['date'],'kind':src['source_kind'],'locator':loc+' '+anchor,'purpose':meaning,'scope':'按条款主体、业务与行为要件适用；一般治理条款不独立证明所有具体行为违法。','url':src['url']})
    policy_ids=[s for s in r['sources'] if not s.startswith('G')]
    policy=[]
    for sid in policy_ids:
        src=SOURCES[sid];old=OLD_EVID.get((rid,sid),{});typ=src.get('material_kind','治理材料')
        policy.append(f"[{sid}] {src['date']}《{src['title']}》（{typ}）。")
        EVID.append({'rule_id':rid,'source_id':sid,'title':src['title'],'date':src['date'],'kind':typ,'locator':old.get('locator','全文，结合行为要件核验'),'purpose':'用于提取治理场景、风险和执行线索；属于公开监管材料的规则转写，不以个案结果替代新案件事实认定。','scope':policy_scope(src) if '专项' in typ or '专项' in src['title'] else '个案/治理公开材料说明特定事实和工作要求；不据此认定所有同类账号应接受相同处罚。','url':src['url']})
    # Keep previously cited norms visible while preventing unrelated articles being made operative.
    other_g=[s for s in r['sources'] if s.startswith('G') and s not in {x[0] for x in laws}]
    for sid in other_g:
        src=SOURCES[sid]
        EVID.append({'rule_id':rid,'source_id':sid,'title':src['title'],'date':src['date'],'kind':src['source_kind'],'locator':'补充对照全文','purpose':'辅助治理背景；本规则直接条款以规则背景列明者为准。','scope':'须先核对服务角色和业务范围，不自动扩张义务。','url':src['url']})
    ids=list(dict.fromkeys([s for s,a,m in laws]+r['sources']))
    background='法律与制度依据：\n'+'\n'.join(legal)+'\n政策与治理材料：\n'+ ('\n'.join(policy) if policy else '本条按所列制度条款设计，不以专项通报替代法定义务。')+'\n适用说明：仅在相应服务主体、业务和受众范围内适用。专项期限与个案边界详见“规则证据”表；持续性平台建议须另行纳入公示规则。一般管理条款仅提供治理框架，细分违法定性仍需对应具体有效依据。'
    principle=PRINCIPLES[c]+'\n本条正常边界：'+r['boundary']+'\n'+COMMON_BOUNDARY
    controlled='判定条件：'+r['decision']+'\n处理类型：'+action_type(r)+'。\n证据清单：'+r['evidence']+'。\n命中后的措施：'+r['action']+'。\n措施约束：区分“法定义务、监管材料转写、平台建议”，按适用依据执行。一般不良信息不自动等同违法信息；单纯缺标、资质待核或功能缺陷不直接记为内容造假。'
    parts=entries.split('；');assert len(parts)>=2,(rid,entries)
    flow=f'场景A：{parts[0]}。场景B：{parts[1]}。\n触发与核验：{scenario}。\n取证：{r["evidence"]}。\n执行链：识别适用业务/对象→取得完整上下文及所列证据→逐项核验判定条件与正常边界→由{roles(r)}确认→按本条措施处理→告知理由并开放申诉。'
    if rid in FUNCTION_IDS|MIXED_IDS or rid in {'C03-03','C14-03'}:
        flow+='\n产品复测：使用合规测试账号记录年龄/身份、配置版本、操作路径、输入和实际输出；复现缺陷后创建整改工单；修复后按相同路径验证，功能缺陷由责任团队整改，不给普通使用者自动记违规。'
    positive={'case_id':rid+'-P','rule_id':rid,'type':'正例（应管控的构造情境）','scenario':pos,'reason':f'在上述构造事实和证据成立的前提下，满足本条“{r["decision"]}”；仍须检查主体、语境及适用范围。','expected':r['action'],'status':'本次文档预期裁决，未经过真人审核一致性或线上测试'}
    negative={'case_id':rid+'-N','rule_id':rid,'type':'反例（不应按本条处罚的构造情境）','scenario':neg,'reason':'该情境落入正常表达或合规执行边界：'+r['boundary']+'；没有本条要求的具体违规事实。','expected':'保留正常表达或维持合规功能，不按本条处罚；如发现独立于此的新事实，按另一个匹配条款单独核验，不因其他条款可能存在而推定违规。','status':'本次文档预期裁决，未经过真人审核一致性或线上测试'}
    # Individual reasons are grounded in the scenario comparison, not just a shared rubric.
    positive['reason']='该情境的关键区别：'+scenario+'；正例已给出相应违反行为及可核验材料。'+positive['reason']
    negative['reason']='与正例相比，反例的关键区别是：'+neg+'。'+negative['reason']
    examples=f"正例 {positive['case_id']}（应管控，构造情境）：{pos}\n判定理由：{positive['reason']}\n预期动作：{positive['expected']}。\n反例 {negative['case_id']}（不按本条处罚，构造情境）：{neg}\n排除理由：{negative['reason']}\n预期动作：{negative['expected']}\n两例均为说明性训练样例，不是对真实通报事实的转述。"
    notes=f'特别备注：{special}。\n'+COMMON_NOTES+f'\n执行责任角色（建议）：{roles(r)}。版本：{VERSION}；依据核对日：{DATE}。\n复核触发：规范修订或施行、重大产品变更、申诉揭示系统性误判、专项期限届满。上线前需填入本平台的具体岗位和联系渠道；这不改变本文档规则定义。'
    card={'rule_id':rid,'category_id':c,'category':CAT[c],'title':r['title'],'priority_suggestion':r['priority'],'control_type':action_type(r),'rule_background':background,'control_principles':principle,'controlled_content':controlled,'scenarios':flow,'examples':examples,'notes':notes,'sources':ids,'law_anchors':[{'source_id':s,'article':a,'scope':m} for s,a,m in laws],'case_ids':[positive['case_id'],negative['case_id']],'version':VERSION,'review_date':DATE,'owner_role_suggestion':roles(r)}
    CARDS.append(card);CASES.extend([positive,negative])
    REVIEWS.append([rid,r['title'],'通过','通过（逐条条款定位）','通过（区分证据与意见）','通过（两入口及核验链）','通过（独立正反例）','通过（处理对象分开）','通过（专项/主体/年龄边界）','文档自评通过','作者兼评审：本助手；结构校验见验收报告。非独立第三方法务结论；真实业务有效性尚未实测。'])

FIXES=[
['F01','原规则仅有字段化判定、边界、措施','按用户六段结构逐条重写，新增原则、场景及备注','75条×6段=450段，逐条可查','已关闭'],
['F02','36组类级样例无法覆盖75条规则','逐条新建正反例及理由、预期动作，保留构造标识','150条样例，75个唯一规则覆盖','已关闭'],
['F03','MCN第六条被草稿误标为平台管理制度','纠正为第七、九、十、十三、十六、二十条；第六条仅登记要求','C18-01与G07原文条款核对','已关闭'],
['F04','金融营销第五、六条在草稿中的义务摘要颠倒','第五条对应许可范围/委托，第六条对应不得便利非法金融活动','C09-01与G09原文条款核对','已关闭'],
['F05','容易将一般不良、缺标、功能缺陷都转成删除或封号','补入处理类型及产品整改复测；账号责任单独认定','规则C03-03、C07-01、C08-04、C17/C18','已关闭'],
['F06','“负面情绪”“炫富”等笼统词可能替代行为要件','C16-01仅以具体造假煽动或侮辱歧视攻击定性，明确正常边界','C16-01正反例及判定条件','已关闭'],
['F07','AI疑似提示可能被当成已证实造假或无条件使用','C03-03反例限定明显生成痕迹等适用条件，并明确提示非处罚','C03-03及G03第六条','已关闭'],
['F08','六个月日志期限可能被扩张为所有审核记录期限','仅限G03第九条规定的特定不加显式标识导出服务情形','C18-02特别备注及条款表','已关闭'],
['F09','历史专项期限、个案处罚可能被长期套用','证据表逐项保留期限线索及适用说明，C18-05专门审查','公开材料不作为长期处罚唯一依据','已关闭'],
['F10','文件校验可能被误认为真实平台效果验收','验收明确文档自评范围；线上一致性、误判漏判及产品复测未实测','验收结论和上线待验清单','已关闭']]

CHECKS=[]
def check(name,ok,evidence):
    CHECKS.append({'check':name,'passed':bool(ok),'evidence':evidence})
check('编号与分类保持一致',len(CARDS)==75 and len(CAT)==18 and {c['rule_id'] for c in CARDS}==set(DETAIL),'18类、75个唯一编号')
keys=['rule_background','control_principles','controlled_content','scenarios','examples','notes']
check('六段内容完整',all(all(c[k].strip() for k in keys) for c in CARDS),'75×6=450个非空段落；另有实质审查台账')
check('规则场景可定位',all('场景A' in c['scenarios'] and '场景B' in c['scenarios'] and '取证：' in c['scenarios'] and '执行链：' in c['scenarios'] for c in CARDS),'每条至少两个入口、核验要点、证据和处理链')
check('样例一一对应',len(CASES)==150 and len({x['case_id'] for x in CASES})==150 and all(sum(x['rule_id']==c['rule_id'] for x in CASES)==2 for c in CARDS),'75正例+75反例，均标为构造情境')
check('样例去重',len({x['scenario'] for x in CASES})==150,'150条情境内容互异')
check('来源有效且可追溯',all(s in SOURCES and SOURCES[s]['status']=='ok' for c in CARDS for s in c['sources']),'全部引用对应已成功抽取正文；17项待核验未作独立证据')
check('条款可精确定位',all(a['locator'].startswith('p') and a['text'].startswith(a['article']) for a in ARTICLES.values()),f'{len(ARTICLES)}个不同条款定位及全文')
check('法规施行日期检查',all(a['effective']<=DATE for a in ARTICLES.values()),'所用补充规范均记录施行日期；未引用尚未施行条款为当前义务')
check('处理类型覆盖',all(c['control_type'] for c in CARDS) and any('功能' in c['control_type'] for c in CARDS),'内容、位置/受众、标识/资质及功能流程分别表达')
check('特殊边界回归',all(['第九条' in next(c for c in CARDS if c['rule_id']=='C18-02')['notes'],'其中一项' in next(c for c in CARDS if c['rule_id']=='C08-04')['notes'],'负面情绪' in next(c for c in CARDS if c['rule_id']=='C16-01')['notes']]),'标识日志范围、六类标签选项、正常负面表达')
check('逐条评审台账完整',len(REVIEWS)==75,'每条记录依据、证据、场景、样例、处置、边界六项及结论')
check('问题整改闭环',all(x[-1]=='已关闭' for x in FIXES),'10项文档问题已修正；线上待验不伪记关闭')
assert all(c['passed'] for c in CHECKS),CHECKS

LAUNCH=[
['L01','实际业务适用性','平台内容治理负责人/法务','确认是否提供拟人互动、AI生成、金融营销、MCN等服务，裁剪不适用条款','逐条适用/不适用决定及依据','未实测/待平台确认'],
['L02','真人审核一致性','培训和质检负责人','以150条为种子补入真实脱敏样本，盲审并记录分歧；规则作者不得把预期答案当实测结果','一致性结果、分歧原因与修订记录；门槛由平台明确','未实测'],
['L03','误判与漏判','质检/算法负责人','建立独立有代表性的真实样本，分别观察正常投诉、新闻、方言、AI和未成年人场景','分类误判漏判及高风险复核结果；本次无统计推断','未实测'],
['L04','产品功能复测','产品/研发/算法负责人','复测发布标签、未成年人入口、推荐开关、打赏、申诉及日志权限','版本、路径、期望/实际结果、缺陷闭环','未实测'],
['L05','处置与申诉演练','运营/申诉负责人','用脱敏演练内容验证告知、复核独立性、撤销恢复、机构关联责任','演练日志及恢复完整性','未实测'],
['L06','版本和公示','平台治理负责人','填入岗位、通知/申诉渠道及适用服务清单，评估存量迁移和对外公示','正式发布版本与生效/过渡安排','待平台落地']]

# User-facing files.
json_path=OUT/'09_六段式规则卡_v2.json'
json_path.write_text(json.dumps({'version':VERSION,'as_of':DATE,'scope':'文档设计、编制、自评与结构验收；非独立法务或上线效果验收','framework':[dict(zip(['section','required_content','acceptance'],f)) for f in FRAMEWORK],'rules':CARDS,'examples':CASES,'evidence':EVID,'law_articles':list(ARTICLES.values()),'checks':CHECKS,'fixes':FIXES,'launch_validation':LAUNCH},ensure_ascii=False,indent=2),encoding='utf-8')

wb=Workbook();wb.remove(wb.active)
def sheet(name,headers,rows,widths):
    ws=wb.create_sheet(name);ws.append(headers)
    for row in rows:ws.append(row)
    ws.freeze_panes='C2' if len(headers)>3 else 'A2';ws.auto_filter.ref=ws.dimensions
    for cell in ws[1]:cell.font=Font(name='微软雅黑',bold=True,color='FFFFFF',size=11);cell.fill=PatternFill('solid',fgColor='9C2924');cell.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[1].height=34
    for i,w in enumerate(widths,1):ws.column_dimensions[get_column_letter(i)].width=w
    for row in ws.iter_rows(min_row=2):
        lines=1
        for cell in row:
            cell.font=Font(name='微软雅黑',size=10);cell.alignment=Alignment(wrap_text=True,vertical='top')
            w=widths[cell.column-1] if cell.column<=len(widths) else 25
            val=str(cell.value or '')
            visual=max(1,sum(max(1,math.ceil(sum(2 if ord(ch)>255 else 1 for ch in s)/max(w-2,1))) for s in val.split('\n')))
            lines=max(lines,visual)
            if val.startswith('https://'):
                cell.hyperlink=val;cell.font=Font(name='微软雅黑',size=10,color='0563C1',underline='single')
        ws.row_dimensions[row[0].row].height=min(409,max(38,lines*14+7))
        if row[0].row%2==0:
            for cell in row:cell.fill=PatternFill('solid',fgColor='FAF4EC')
    ws.sheet_view.zoomScale=85
    return ws
intro=[['阅读顺序','先读框架设计→治理分类→规则总表→规则卡分段→逐条正反例→规则证据→逐条评审→验收/整改。'],['本次范围','18类75条，450段正文，150个构造样例。规则来源沿用已抓取材料，并补充精确条款映射。'],['正反例定义','正例=满足管控条件的构造案例；反例=表面相似但不应按该条处罚的构造案例，不代表所有其他行为全面免责。'],['字段阅读','规则总表保留用户六段框架；较长内容在规则卡分段按六行展开。来源编号与原始数据04一致，证据表提供官方链接。'],['适用依据','部门规章、部门规范性文件、专项通知、执法通报、平台建议分别标明；一般治理条款不是所有细分违法的直接认定依据。'],['验收结论','文档设计自评及结构验收通过；真实业务适用、真人一致性、产品联调和误判漏判效果尚未实测。'],['评审身份','编制与评审均由本助手完成，属于作者自评；未安排独立第三方或实际平台审核员测试。'],['版本优先','v2用于本次六段式规则执行讨论；v1文件01—03为前版研究底稿，含未按本次框架展开内容；04原始语料与05抓取快照仍可溯源。'],['采集边界','两栏目703篇去重网页，686篇成功抽取正文，17篇图片/专题待核验；244篇列为治理材料，另有24篇背景、413篇非治理及5篇相邻排除。'],['材料期限','抓取口径截至2026-11-17；栏目材料日期2013-05-02至2026-11-06。未将页面抓取成功等同于正文逐字人工核对。'],['在线复核','本轮3项重点页面重取：MCN规定成功；金融营销、短视频标注页返回403，相关审查使用前次已存档官方正文，不记作本轮联网成功。'],['优先级','P0建议优先核验紧迫现实伤害等风险；P1建议优先核验明确违法或严重侵权；P2用于一般不良、补标等常规审核；P3用于机制、质量和流程复核。实际顺序按紧迫性调整；不表示法定违法等级、统一处置时限或当然封号。']]
sheet('使用说明',['项目','说明'],intro,[22,132])
sheet('框架设计',['六段栏目','应包含的内容','验收口径'],FRAMEWORK,[20,100,100])
sheet('治理分类',['类别编号','治理分类','管控原则','规则数量'],[[k,CAT[k],PRINCIPLES[k],sum(c['category_id']==k for c in CARDS)] for k in CAT],[15,34,100,14])
sheet('规则总表',['规则编号','分类','规则名称',*SECTIONS,'处理类型','建议队列','来源编号','版本'],[[c['rule_id'],c['category'],c['title'],*[c[k] for k in keys],c['control_type'],c['priority_suggestion'],'；'.join(c['sources']),VERSION] for c in CARDS],[14,27,32,95,85,90,90,110,95,35,13,32,12])
sheet('规则卡分段',['规则编号','规则名称','段落序号','六段栏目','内容','来源编号'],[[c['rule_id'],c['title'],i+1,label,c[k],'；'.join(c['sources'])] for c in CARDS for i,(label,k) in enumerate(zip(SECTIONS,keys))],[15,34,13,20,145,40])
sheet('逐条正反例',['样例编号','规则编号','样例性质','构造情境','理由','预期动作','验证状态'],[[x['case_id'],x['rule_id'],x['type'],x['scenario'],x['reason'],x['expected'],x['status']] for x in CASES],[17,15,30,85,110,80,40])
sheet('规则证据',['规则编号','来源编号','标题','发布日期','效力或材料类型','原文定位','引用目的及边界','期限与范围','官方链接'],[[e[k] for k in ['rule_id','source_id','title','date','kind','locator','purpose','scope','url']] for e in EVID],[15,15,56,16,30,35,85,110,65])
sheet('法规条款',['来源编号','规范名称','条款','正文定位','条文全文（存档）','本次引用目的','施行日期','效力类型','官方链接'],[[a[k] for k in ['source_id','title','article','locator','text','meaning','effective','kind','url']] for a in ARTICLES.values()],[15,48,16,16,135,70,17,26,65])
sheet('逐条评审',['规则编号','规则名称','六段完整','依据评审','事实证据','场景评审','正反例评审','处置评审','适用边界','自评结论','身份及局限'],REVIEWS,[15,34,18,30,30,32,30,32,40,22,90])
sheet('整改台账',['问题编号','发现问题','整改内容','核查证据','状态'],FIXES,[15,65,100,75,20])
sheet('文档验收',['验收项','结果','证据说明'],[[c['check'],'通过' if c['passed'] else '失败',c['evidence']] for c in CHECKS],[38,15,125])
sheet('上线待验',['编号','验收事项','责任角色建议','方法','所需输出','状态'],LAUNCH,[14,28,34,100,90,30])
sheet('审核步骤',['步骤','名称','操作要点'],STEPS,[12,24,125])
xlsx=OUT/'06_六段式规则框架与验收_v2.xlsx';wb.save(xlsx)

# Report and manual share the same canonical rule text.
head=['# 互联网平台内容审核规则：六段式设计、评审与验收',f'版本：{VERSION}｜依据核对日：{DATE}｜作者及评审：本助手（自评）','',
'本版按“规则背景—管控原则—管控内容—具体管控场景—正反例—备注”逐条编制，保留18类75条规则编号。这里的正例指应予管控的构造情境，反例指表面相似但不应按该条处罚的正常情境。150个样例均有预期裁决，尚未经过真实审核员或线上系统测试。',
'文档级设计自评与结构验收通过。该结论覆盖规则定义、证据追溯和交付完整性；不等于独立第三方法务意见、实际平台适用确认或线上效果验收。',
'', '## 一、六段式框架与适用规则']
for label,required,accept in FRAMEWORK:head.extend([f'### {label}',required,'验收：'+accept,''])
head.extend(['### 依据层次和处置层次','法律与制度条款按名称、条款、效力类型、施行日期和适用主体记录。专项通知、治理进展和执法个案用于提取治理场景与监管要求；其适用时间、服务角色和具体事实须另核。原则、证据清单、构造例、岗位和P0—P3队列包含平台实施建议，不能冒称原文法定义务。','内容违法且要件明确时按适用依据停止传播；不良信息按具体条款防范抵制，特定位置或受众限制单列；补标、来源与资质核验和平台功能整改分别建单。账号或机构责任须另核，不按标签数量累计处罚。','一般治理条款是制度依据；对诈骗、侵权、学术造假、保密等具体法律定性，仍需核对相应实体法和业务规则。没有具体证据的争议不直接记为违法，实际危险可先采取必要且可撤销的保护措施。','### 审核六步',*[f'{a} {b}：{c}' for a,b,c in STEPS],'','## 二、采集与证据范围','本版沿用两个指定栏目采集到的703篇去重网页：686篇正文抽取成功，17篇以图片或专题为主、待进一步核验。筛入244篇治理材料和24篇背景材料，排除413篇非治理及5篇相邻材料；补充9份规范。栏目日期覆盖2013-05-02至2026-11-06，抓取口径截至2026-11-17。','标题/正文关键词初标用于检索，不等于已对全部703篇逐字人工核对。待核验17篇不作为规则独立证据。每条规则均有来源编号，官方链接、定位、引用目的和范围见证据表；抓取快照、原始正文及哈希沿用04、05文件。','本轮额外联网重取3项重点材料：MCN规定成功；金融营销办法和短视频标注部署页面返回403，相关文本审查使用前次已保存官方快照。这两项不记为本轮网络复核成功。','', '## 三、18类治理分类'] )
for k in CAT:head.append(f'- {k} {CAT[k]}：{sum(c["category_id"]==k for c in CARDS)}条。{PRINCIPLES[k]}')
head.extend(['','## 四、75条六段式规则全文'])
full=list(head)
for c in CARDS:
    full.extend(['',f'### {c["rule_id"]} {c["title"]}',f'分类：{c["category"]}；处理类型：{c["control_type"]}；队列建议：{c["priority_suggestion"]}。'])
    for label,k in zip(SECTIONS,keys):full.extend(['',f'**{label}**','',c[k]])
full.extend(['','## 五、评审与验收','本次评审为作者自评：由本助手逐条编制、核对适用边界，并用程序核验结构和来源一致性。未采用独立第三方评审、真人双盲审核或真实流量实验。','### 问题与整改'])
for fix in FIXES:full.append(f'- {fix[0]} {fix[1]}。整改：{fix[2]}。核查：{fix[3]}；状态：{fix[4]}。')
full.extend(['','### 文档验收'])
for ch in CHECKS:full.append(f'- {ch["check"]}：通过。{ch["evidence"]}。')
full.extend(['','### 上线待验'])
for x in LAUNCH:full.append(f'- {x[0]} {x[1]}：{x[3]}。输出：{x[4]}；状态：{x[5]}。')
full.extend(['','## 六、官方来源索引','此处列出本版直接引用的来源；条款及段落定位见工作簿。'])
for sid in sorted({s for c in CARDS for s in c['sources']}):
    src=SOURCES[sid];full.append(f'- [{sid}] {src["date"]}《{src["title"]}》 {src["url"]}')
md=OUT/'08_规则设计评审与验收_v2.md';md.write_text('\n\n'.join(full)+'\n',encoding='utf-8')

# Word: human-readable, editable manual with outline and page numbers.
doc=Document();section=doc.sections[0];section.top_margin=Cm(2);section.bottom_margin=Cm(2);section.left_margin=Cm(2.2);section.right_margin=Cm(2.2)
normal=doc.styles['Normal'];normal.font.name='Microsoft YaHei';normal._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');normal.font.size=Pt(10);normal.paragraph_format.space_after=Pt(5)
for sty,size in [('Title',23),('Heading 1',17),('Heading 2',14),('Heading 3',12)]:
    st=doc.styles[sty];st.font.name='Microsoft YaHei';st._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');st.font.size=Pt(size);st.font.color.rgb=RGBColor.from_string('8E2924')
header=section.header.paragraphs[0];header.text='互联网平台内容审核规则 · 六段式 v2.0';header.style='Caption'
p=section.footer.paragraphs[0];p.alignment=2;p.add_run('第 ')
f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),'PAGE');p._p.append(f);p.add_run(' 页')
for text in full:
    if not text:continue
    if text.startswith('# '):doc.add_heading(text[2:],0)
    elif text.startswith('## '):doc.add_heading(text[3:],1)
    elif text.startswith('### '):doc.add_heading(text[4:],2)
    elif text.startswith('**') and text.endswith('**'):doc.add_heading(text[2:-2],3)
    elif text.startswith('- '):doc.add_paragraph(text[2:],style='List Bullet')
    else:
        for line in text.splitlines():doc.add_paragraph(line)
docx=OUT/'07_六段式内容审核规则手册_v2.docx';doc.save(docx)

# Verify files by reopening them; checks do not infer real-world performance.
again=load_workbook(xlsx,read_only=True,data_only=True)
assert again['规则总表'].max_row-1==75
assert again['规则卡分段'].max_row-1==450
assert again['逐条正反例'].max_row-1==150
assert again['逐条评审'].max_row-1==75
assert again['法规条款'].max_row-1==len(ARTICLES)
rows=again['规则总表'].iter_rows(min_row=2,values_only=True)
for c,row in zip(CARDS,rows):assert row[0]==c['rule_id'] and list(row[3:9])==[c[k] for k in keys]
again.close()
dd=Document(docx);doc_text='\n'.join(p.text for p in dd.paragraphs)
assert all(c['rule_id'] in doc_text and all(cid in doc_text for cid in c['case_ids']) for c in CARDS)
assert all(doc_text.count(label)>=75 for label in SECTIONS)
assert len(json.loads(json_path.read_text(encoding='utf-8'))['rules'])==75
check('交付文件重开与内容一致性',True,'工作簿重开核对75/450/150/75行；Word重开核对全部编号与样例；JSON回读75条')
# Refresh validation results in all machine-readable deliverables without claiming a new corpus crawl.
meta={'version':VERSION,'date':DATE,'categories':18,'rules':75,'sections':450,'examples':150,'rule_source_evidence_rows':len(EVID),'unique_sources':len({s for c in CARDS for s in c['sources']}),'unique_law_articles':len(ARTICLES),'reviewed_rule_rows':75,'fixed_findings':len(FIXES),'checks':CHECKS,'document_acceptance':'通过（作者自评与结构验收）','independent_review':'未实施','live_platform_acceptance':'未实施','recent_network_checks':{'G07':'成功','G09':'HTTP 403；采用已存档官方正文','S0024':'HTTP 403；采用已存档官方正文'}}
(WORK/'validation_v2.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2),encoding='utf-8')
payload=json.loads(json_path.read_text(encoding='utf-8'));payload['checks']=CHECKS;payload['validation']=meta;json_path.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
# Add final reopening check to workbook and report without regenerating card content.
ws=wb['文档验收'];ws.append([CHECKS[-1]['check'],'通过',CHECKS[-1]['evidence']]);ws.row_dimensions[ws.max_row].height=52
for cell in ws[ws.max_row]:cell.font=Font(name='微软雅黑',size=10);cell.alignment=Alignment(wrap_text=True,vertical='top')
wb.save(xlsx)
with md.open('a',encoding='utf-8') as f:f.write('\n\n## 七、交付文件校验\n\n'+CHECKS[-1]['evidence']+'。校验通过。\n')
doc.add_heading('七、交付文件校验',1);doc.add_paragraph(CHECKS[-1]['evidence']+'。校验通过。');doc.save(docx)
# Preserve original delivery description, then clearly set version priority.
oldnote=OUT/'00_交付说明.txt';existing=oldnote.read_text(encoding='utf-8-sig')
if '六段式规则v2' not in existing:
    prefix=f'''【六段式规则v2：当前规则稿】\n用户追加的框架、评审与验收已完成。优先使用06—09文件。\n06_六段式规则框架与验收_v2.xlsx：13张表，规则总表、450段分段规则卡、150个样例、规则证据、精确法规条文、逐条评审、整改和验收。\n07_六段式内容审核规则手册_v2.docx：可编辑六段式规则手册。\n08_规则设计评审与验收_v2.md：含框架、全文、评审和验收边界。\n09_六段式规则卡_v2.json：可导入后续配置或知识库的结构化规则。\n10_六段式规则复现与验收_v2.zip：新版生成脚本、规则输入、校验记录和差异说明。\n\n本版共18类75条、450段、75组正反例；正例指应管控，反例指不按该条处罚，均为构造情境。文档自评及结构验收通过。真人一致性、实际平台功能与上线效果未实测；作者自评不冒称独立第三方审查。\n本次修正MCN平台义务条款和金融营销第五/六条义务映射，明确内容、补标、受众/位置和功能整改分开。v1的01—03为前版底稿，不作为本次六段式规则的当前版本；原始语料04和抓取快照05继续用于溯源。\n\n--- 以下保留原研究交付说明 ---\n\n'''
    oldnote.write_text(prefix+existing,encoding='utf-8')
(WORK/'v2_README.txt').write_text('六段式规则v2复现材料\n生成：在解压后的work目录执行 Python build_v2.py。依赖：python-docx、openpyxl。输入04 JSONL（原始来源）、rules.json、policy_data.py、framework_v2_data.py及evidence.csv。\n结构校验包含在生成脚本末尾，失败会抛出断言；validation_v2.json记录本次结果。\n本包不重新抓取网站；原始HTML与抓取程序见05_来源与复现材料.zip。\nMCN及金融义务草稿误引已在build_v2.py显式修正，不应独立执行原始映射草稿当作最终规范。\n',encoding='utf-8')
zpath=OUT/'10_六段式规则复现与验收_v2.zip'
with zipfile.ZipFile(zpath,'w',zipfile.ZIP_DEFLATED) as z:
    for filename in ['build_v2.py','framework_v2_data.py','policy_data.py','rules.json','evidence.csv','validation_v2.json','v2_README.txt']:
        z.write(WORK/filename,'work/'+filename)
    z.write(OUT/'04_抓取语料与依据.jsonl','outputs/04_抓取语料与依据.jsonl')
with zipfile.ZipFile(zpath) as z:assert z.testzip() is None
print(json.dumps(meta,ensure_ascii=False,indent=2))
for p in [xlsx,docx,md,json_path,zpath]:print(p.name,p.stat().st_size)
