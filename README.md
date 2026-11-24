# 中国网信网不良内容治理通报与审核规则

本项目抓取中国网信网以下两个栏目，并据通报整理互联网平台内容审核规则：

- 互联网内容管理：`A093704`
- “清朗”系列专项行动：`A093711`

规则卡按“规则背景、管控原则、管控内容、具体管控场景、正反例、备注”六段式组织，当前规则输入和生成结果均保留在仓库中。

## 自动运行

GitHub Actions 每日北京时间 01:20（UTC 前一日 17:20）运行，也支持手动触发。手动运行时可以传入 `scrape_date`（YYYY-MM-DD）；不传则使用运行环境的北京时间日期。工作流会：

1. 安装 Python 依赖；
2. 抓取列表、分页和正文；
3. 生成 v1 和六段式 v2 交付物；
4. 运行结构校验；
5. 将原始 HTML、列表响应和完整 outputs 上传为 artifact；
6. 仅在生成文件变化时提交并推送到默认分支。

原始 HTML 和列表响应被 .gitignore 排除，以避免仓库随每次运行持续膨胀；它们仍会作为带日期的 Actions artifact 保存。

## 本地运行

需要 Python 3.11+：

```powershell
python -m pip install -r requirements.txt
$env:SCRAPE_CUTOFF = (Get-Date).ToString('yyyy-MM-dd')
$env:SCRAPE_DATE = $env:SCRAPE_CUTOFF
python work/collect.py
python work/build.py
python work/build_v2.py
python work/validate_outputs.py
```

如需复现历史快照，可显式设置 `SCRAPE_CUTOFF` 和 `SCRAPE_DATE`。

## 主要产物

- `outputs/06_六段式规则框架与验收_v2.xlsx`
- `outputs/07_六段式内容审核规则手册_v2.docx`
- `outputs/08_规则设计评审与验收_v2.md`
- `outputs/09_六段式规则卡_v2.json`
- `outputs/10_六段式规则复现与验收_v2.zip`
- `work/rules.json`：规则输入
- `work/source_id_registry.json`：稳定来源编号注册表

## 范围与限制

采集内容来自中国网信网公开页面；抓取结果和规则整理用于研究、产品规则设计和审核流程讨论，不替代法律意见、监管解释或平台上线前的法务审查。真实审核员一致性、模型效果、误杀漏放率和投诉处置效果仍需要在目标平台上单独测试和验收。
