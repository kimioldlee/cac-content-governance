六段式规则v2复现材料
生成：在解压后的work目录执行 Python build_v2.py。依赖：python-docx、openpyxl。输入04 JSONL（原始来源）、rules.json、policy_data.py、framework_v2_data.py及evidence.csv。
结构校验包含在生成脚本末尾，失败会抛出断言；validation_v2.json记录本次结果。
本包不重新抓取网站；原始HTML与抓取程序见05_来源与复现材料.zip。
MCN及金融义务草稿误引已在build_v2.py显式修正，不应独立执行原始映射草稿当作最终规范。
