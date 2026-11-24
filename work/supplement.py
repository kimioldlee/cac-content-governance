import json,re
from collect import article,WORK,fetch,text
SPECS=[
('G01','网络信息内容生态治理规定','https://www.cac.gov.cn/2019-12/20/c_1578375159509309.htm','2020-03-01','部门规章'),
('G02','可能影响未成年人身心健康的网络信息分类办法','https://www.cac.gov.cn/2026-01/23/c_1770728781060093.htm','2026-03-01','部门规范性文件'),
('G03','人工智能生成合成内容标识办法','https://www.cac.gov.cn/2025-03/14/c_1743654684782215.htm','2025-09-01','部门规范性文件'),
('G04','网络暴力信息治理规定','https://www.cac.gov.cn/2024-06/14/c_1720043894161555.htm','2024-08-01','部门规章'),
('G05','互联网信息服务算法推荐管理规定','https://www.cac.gov.cn/2022-01/04/c_1642894606364259.htm','2022-03-01','部门规章'),
('G06','人工智能拟人化互动服务管理暂行办法','https://www.cac.gov.cn/2026-04/10/c_1777558395078289.htm','2026-07-15','部门规章'),
('G07','互联网信息内容多渠道分发服务管理规定','https://www.cac.gov.cn/2026-05/29/c_1781795864412597.htm','2026-09-01','部门规章'),
('G08','互联网军事信息传播管理办法','https://www.cac.gov.cn/2025-02/08/c_1740714184323988.htm','2025-03-01','部门规范性文件'),
('G09','金融产品网络营销管理办法','https://www.cac.gov.cn/2026-04/24/c_1778769008779432.htm','2026-09-30','部门规范性文件')]
r=[]
for id,title,url,effective,kind in SPECS:
    x=article({'id':id,'title':title,'url':url,'date':re.search(r'/(\d{4}-\d{2})/(\d{2})/',url).group(1)+'-'+re.search(r'/(\d{4}-\d{2})/(\d{2})/',url).group(2),'channels':['补充规范依据'],'list_pages':[]})
    x['effective']=effective;x['source_kind']=kind;r.append(x)
    print(id,x['status'],x['chars'],title,'\n',x['text'][:160]+' ... '+x['text'][-400:],flush=True)
(WORK/'supplement.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')

