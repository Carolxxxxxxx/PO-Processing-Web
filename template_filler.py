import sys
import pdfplumber
import openpyxl
import re
import os

# **获取命令行参数**（确保正确解析上传的 `PO` 和 `价格表`）
pdf_path = sys.argv[1] if len(sys.argv) > 1 else None
price_path = sys.argv[2] if len(sys.argv) > 2 else None

if not pdf_path or not price_path:
    print("❌ 错误：未提供 `PO.pdf` 或 `价格表.xlsx`，请检查上传！")
    sys.exit(1)  # **终止程序**

# ========== **解析 PO 并提取数据** ==========
def extract_data_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()
        if not text:
            print("❌ 无法提取文本，可能是扫描版 PDF。")
            return None, None, None

        part_numbers = []
        ordered_quantities = []
        po_number = None

        # 提取 PO 号
        po_match = re.search(r"Order Number\s*([\d-]+)", text)
        if po_match:
            po_number = po_match.group(1)

        # 解析 Part Number 和 订单数量
        lines = text.split("\n")
        for i in range(len(lines)):
            match = re.search(r"(BHB\d{3,}-CLRK|BHW\d{3,}-CLRK)", lines[i])
            if match:
                part_number = match.group(1)
                ordered_match = re.findall(r"(\d{2,}\.00)", lines[i])
                ordered_quantity = int(float(ordered_match[-1])) if ordered_match else "解析错误"

                part_numbers.append(part_number)
                ordered_quantities.append(ordered_quantity)

        return part_numbers, ordered_quantities, po_number

# ========== **读取价格表** ==========
def load_price_list(price_path):
    wb = openpyxl.load_workbook(price_path)
    ws = wb.active
    price_dict = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            part_number = row[0]
            price = row[2] if isinstance(row[2], (int, float)) else "N/A"
            units_per_case = 250 if "Case-250" in row[1] else 200
            price_dict[part_number] = (price, units_per_case)

    return price_dict

# ========== **生成 INVOICE 和 PACKING LIST** ==========
part_numbers, ordered_quantities, po_number = extract_data_from_pdf(pdf_path)
if part_numbers and ordered_quantities and po_number:
    price_list = load_price_list(price_path)
    print(f"✅ 解析 PO 成功，发现 {len(part_numbers)} 个产品。")
else:
    print("❌ PO 解析失败，请检查文件格式。")
    sys.exit(1)  # **终止程序**
