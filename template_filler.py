import sys
import pdfplumber
import openpyxl
import os
import re

# 📌 确保命令行参数正确
if len(sys.argv) < 3:
    print("❌ 错误：未提供 `PO.pdf` 或 `价格表.xlsx`，请检查上传！")
    sys.exit(1)

pdf_path = sys.argv[1]
price_path = sys.argv[2]

# ========== 1️⃣ 解析 PDF 提取 PO 号、Part Number、箱数 ==========
def extract_data_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        print(f"❌ PO 文件 `{pdf_path}` 未找到！")
        return None, None, None

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()
        if not text:
            print("❌ 无法提取文本，可能是扫描版 PDF。")
            return None, None, None

        part_numbers = []
        ordered_quantities = []
        po_number = None

        po_match = re.search(r"Order Number\s*([\d-]+)", text)
        if po_match:
            po_number = po_match.group(1)

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

# ========== 2️⃣ 读取价格表 ==========
def load_price_list(price_path):
    if not os.path.exists(price_path):
        print(f"❌ 价格表 `{price_path}` 未找到！")
        return {}

    wb = openpyxl.load_workbook(price_path)
    ws = wb.active
    price_dict = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            part_number = row[0]
            price = row[2] if isinstance(row[2], (int, float)) else "N/A"
            units_per_case = 250 if "Case-250" in row[1] else 200
            nw = row[3] if isinstance(row[3], (int, float)) else "N/A"
            gw = row[4] if isinstance(row[4], (int, float)) else "N/A"
            price_dict[part_number] = (price, units_per_case, nw, gw)

    return price_dict

# ========== 3️⃣ 生成 INVOICE ==========
def fill_invoice(template_path, output_path, part_numbers, ordered_quantities, po_number, price_list):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    ws["J9"] = po_number

    start_row = 15
    for i, (part_number, ordered_quantity) in enumerate(zip(part_numbers, ordered_quantities)):
        row = start_row + i
        price, units_per_case, _, _ = price_list.get(part_number, ("N/A", 250, "N/A", "N/A"))
        product_quantity = ordered_quantity * units_per_case if isinstance(ordered_quantity, int) else "N/A"

        ws[f"B{row}"] = part_number
        ws[f"C{row}"] = ordered_quantity
        ws[f"E{row}"] = product_quantity
        ws[f"H{row}"] = price

    wb.save(output_path)
    print(f"✅ INVOICE 生成成功：{output_path}")

# ========== 4️⃣ 生成 PACKING LIST ==========
def fill_packing_list(template_path, output_path, part_numbers, ordered_quantities, po_number, price_list):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    ws["K11"] = po_number

    start_row = 17
    for i, (part_number, ordered_quantity) in enumerate(zip(part_numbers, ordered_quantities)):
        row = start_row + i
        price, units_per_case, nw, gw = price_list.get(part_number, ("N/A", 250, "N/A", "N/A"))

        ws[f"A{row}"] = part_number
        ws[f"D{row}"] = f"=F{row} * P{row}"  # **Excel 公式**
        ws[f"F{row}"] = ordered_quantity
        ws[f"H{row}"] = f"=F{row} * N{row}"  # **Excel 公式**
        ws[f"N{row}"] = price
        ws[f"O{row}"] = nw
        ws[f"P{row}"] = units_per_case

    wb.save(output_path)
    print(f"✅ PACKING LIST 生成成功：{output_path}")

# ========== 5️⃣ 运行主程序 ==========
part_numbers, ordered_quantities, po_number = extract_data_from_pdf(pdf_path)
if not part_numbers or not ordered_quantities or not po_number:
    print("❌ PO 解析失败，请检查文件格式。")
    sys.exit(1)

price_list = load_price_list(price_path)
if not price_list:
    print("❌ 价格表解析失败，请检查文件格式。")
    sys.exit(1)

invoice_output = f"INVOICE_{po_number}.xlsx"
packing_list_output = f"PACKING_LIST_{po_number}.xlsx"

fill_invoice("INVOICE.xlsx", invoice_output, part_numbers, ordered_quantities, po_number, price_list)
fill_packing_list("PACKING_LIST.xlsx", packing_list_output, part_numbers, ordered_quantities, po_number, price_list)

print("🎉 所有文件生成完毕！")
