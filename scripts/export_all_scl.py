import sys
import io
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

with open('database/norms_database.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

for code in ['E6C.3.05', 'E6C.3.01', 'E6C.3.02', 'E6C.3.03', 'E6C.3.04', 'E6C.3.06', 'E6C.1.31', 'E6B.1.13.01']:
    item = data.get(code, {})
    print('==============================')
    print('CODE:', code, '| NAME:', item.get('name'))
    print('UNIT:', item.get('unit'), '| TOTAL NC:', item.get('total_labor_steps'))
    print('LEGAL BASIS:', item.get('legal_basis'))
    print('STEPS:')
    for s in item.get('work_steps', []):
        try:
            pct_val = float(s.get('dvt_step', 0)) * 100
            pct_str = f"{pct_val:.1f}%"
        except:
            pct_str = str(s.get('dvt_step'))
        print(f"  Step {s['step']}: {s['name']} | {pct_str} | {s.get('rate_calculated')} cong")
    print('MACHINES:')
    for m in item.get('resources', {}).get('machines', []):
        print(f"  - {m.get('name')}: {m.get('rate')} {m.get('unit')}")
