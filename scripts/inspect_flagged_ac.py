import openpyxl, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook(r'C:\Users\Admin\Downloads\DANH MUC CONG TAC_da tra ma.xlsx', data_only=True)
ws = wb.active

categories = defaultdict(list)

for r in range(6, ws.max_row + 1):
    val_ac = ws.cell(row=r, column=29).value
    if val_ac is not None and str(val_ac).strip() in ['1', '1.0', 1]:
        item = {
            'row': r,
            'stt': ws.cell(row=r, column=1).value,
            'name': ws.cell(row=r, column=2).value,
            'loc': ws.cell(row=r, column=3).value,
            'desc': ws.cell(row=r, column=4).value,
            'qty': ws.cell(row=r, column=5).value,
            'unit': ws.cell(row=r, column=6).value,
            'ma_tk': ws.cell(row=r, column=7).value,
            'ma_khvt': ws.cell(row=r, column=9).value,
            'ma_ktat': ws.cell(row=r, column=18).value,
            'ma_tu': ws.cell(row=r, column=27).value,
            'ly_do': ws.cell(row=r, column=28).value,
            'ghi_chu': ws.cell(row=r, column=26).value
        }
        categories[item['ma_tu']].append(item)

print('Total items with Col AC = 1:', sum(len(v) for v in categories.values()))
print('Number of distinct ma_tu groups:', len(categories))
for k, items_list in sorted(categories.items(), key=lambda x: len(x[1]), reverse=True):
    print(f"\n=== Group: {k} (Count: {len(items_list)}) ===")
    for it in items_list[:3]:
        name_short = str(it['name'])[:70]
        ly_do_short = str(it['ly_do'])[:90]
        print(f"  Row {it['row']} (STT {it['stt']}): {name_short} | Qty={it['qty']} {it['unit']} | TK={it['ma_tk']} | KHVT={it['ma_khvt']} | KTAT={it['ma_ktat']}")
        print(f"    Ly do: {ly_do_short}")
    if len(items_list) > 3:
        print(f"  ... and {len(items_list)-3} more items")
