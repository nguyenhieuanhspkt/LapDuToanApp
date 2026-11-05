import fitz
import sys

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\PL QĐ Chuyển đổi NC SCL thiết bị NMNĐ đốt than công suất tổ máy 600MW.pdf'

doc = fitz.open(pdf_path)
print(f"Total pages: {len(doc)}")

output_txt = 'pdf_extracted_text.txt'
with open(output_txt, 'w', encoding='utf-8') as out:
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        out.write(f"\n--- PAGE {page_num + 1} ---\n")
        out.write(text)

print(f"Extracted all text to {output_txt}")
