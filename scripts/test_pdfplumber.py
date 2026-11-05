import pdfplumber

pdf_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Ca máy\ca máy 21 LĐ.pdf'

with pdfplumber.open(pdf_path) as pdf:
    for pno in [9, 10]:
        print(f"=== PAGE {pno+1} ===")
        page = pdf.pages[pno]
        tables = page.extract_tables()
        print(f"Tables extracted: {len(tables)}")
        if tables:
            for t in tables:
                print(f"Table rows: {len(t)}")
                for row in t[:6]:
                    print(" | ".join([str(c or '').replace('\n', ' ') for c in row]))
