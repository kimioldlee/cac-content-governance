"""Validate the generated corpus and six-section rule deliverables."""
import json, os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORK = ROOT / 'work'
OUT = ROOT / 'outputs'
DATE = os.environ.get('SCRAPE_DATE') or __import__('datetime').date.today().isoformat()

def read_json(path):
    return json.loads(path.read_text(encoding='utf-8'))

required = [
    OUT / '04_抓取语料与依据.jsonl',
    OUT / '06_六段式规则框架与验收_v2.xlsx',
    OUT / '07_六段式内容审核规则手册_v2.docx',
    OUT / '08_规则设计评审与验收_v2.md',
    OUT / '09_六段式规则卡_v2.json',
    OUT / '10_六段式规则复现与验收_v2.zip',
    WORK / 'corpus.json',
    WORK / 'rules.json',
    WORK / 'source_id_registry.json',
]
for path in required:
    assert path.exists() and path.stat().st_size > 0, f'missing or empty: {path}'

corpus = read_json(WORK / 'corpus.json')
assert corpus and len({row['url'] for row in corpus}) == len(corpus)
assert all(row.get('date', '') <= DATE for row in corpus), 'corpus contains a notice after SCRAPE_DATE'

rules = read_json(WORK / 'rules.json')
assert len(rules) >= 1
rule_ids = {row['id'] for row in rules}
assert len(rule_ids) == len(rules)
assert all(row.get('decision') and row.get('action') and row.get('sources') for row in rules)

payload = read_json(OUT / '09_六段式规则卡_v2.json')
assert payload.get('as_of') == DATE, (payload.get('as_of'), DATE)
cards = payload.get('rules', [])
assert {row['rule_id'] for row in cards} == rule_ids
sections = {'rule_background', 'control_principles', 'controlled_content', 'scenarios', 'examples', 'notes'}
for card in cards:
    missing = sections - set(card)
    assert not missing, f"{card.get('rule_id')} missing {sorted(missing)}"
    assert all(card.get(key) not in (None, '') for key in sections), f"{card.get('rule_id')} has empty six-section field"
assert payload.get('validation', {}).get('document_acceptance') is not None
print(json.dumps({'date': DATE, 'corpus_records': len(corpus), 'rules': len(rules), 'rule_cards': len(cards), 'status': 'ok'}, ensure_ascii=False, indent=2))
