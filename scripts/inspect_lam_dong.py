import fitz
import openpyxl
import os
import sys

folder = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng'

print("Inspecting folder:", folder)
for f in os.listdir(folder):
    p = os.path.join(folder, f)
    print(f" - {f}: {os.path.getsize(p)} bytes")

# 1. 5898.pdf
p_5898 = os.path.join(folder, '5898.pdf')
if os.path.exists(p_5898):
    doc = fitz.open(p_5898)
    print(f"\n=== 5898.pdf ({len(doc)} pages) ===")
    for i in range(min(5, len(doc))):
        print(f"--- PAGE {i+1} ---")
        print(doc[i].get_text()[:800])

# 2. ca máy 21 LĐ.pdf
p_camay = os.path.join(folder, 'ca máy 21 LĐ.pdf')
if os.path.exists(p_camay):
    doc = fitz.open(p_camay)
    print(f"\n=== ca máy 21 LĐ.pdf ({len(doc)} pages) ===")
    for i in range(min(3, len(doc))):
        print(f"--- PAGE {i+1} ---")
        print(doc[i].get_text()[:800])

# 3. Gia_VLXD_LamDong_T08_2026.xlsx
p_vlxd = os.path.join(folder, 'Gia_VLXD_LamDong_T08_2026.xlsx')
if os.path.exists(p_vlxd):
    wb = openpyxl.load_workbook(p_vlxd, read_only=True)
    print(f"\n=== Gia_VLXD_LamDong_T08_2026.xlsx (Sheets: {wb.sheetnames}) ===")
    for s in wb.sheetnames:
        ws = wb[s]
        print(f"Sheet {s}: max_row={ws.max_row}, max_col={ws.max_column}")

# 4. Đơn giá mới nhất.docx
p_doc = os.path.join(folder, 'Đơn giá mới nhất.docx')
if os.path.exists(p_doc):
    import zipfile, xml.etree.ElementTree as ET
    with zipfile.ZipFile(p_doc, 'r') as z:
        xml_content = z.read('word/document.xml')
        tree = ET.fromstring(xml_content)
        texts = [elem.text for elem in tree.iter() if elem.text]
        print(f"\n=== Đơn giá mới nhất.docx (len text: {len(texts)}) ===")
        print(''.join(texts[:500]))
