import json
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

with open('database/norms_database.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

scl_codes = ['E6C.3.01', 'E6C.3.02', 'E6C.3.03', 'E6C.3.04', 'E6C.3.05', 'E6C.3.06', 'E6C.2.04', 'E6C.1.23', 'E6C.1.31', 'E6B.1.13.01', 'E6C.1.18', 'E6BOP.1.03', 'E6BOP.1.01']

for k in scl_codes:
    if k in data:
        v = data[k]
        print(f"*** {k}: {v.get('name')} | Total: {v.get('total_labor_steps')} | unit: {v.get('unit')} | basis: {v.get('legal_basis')}")
        for s in v.get('work_steps', []):
            print(f"   {s['step']}. {s['name']} | %: {s.get('dvt_step')} | rate: {s.get('rate_calculated')}")
        print("   Resources:", v.get('resources', {}).get('machines', []))
    else:
        print(f"*** {k}: NOT IN DB")
    print()
