
cd /Users/carol/Desktop/POPICI
nano template_filler.pyimport pdfplumber
import openpyxl
import re
import os
import sys

# ========== 1. 解析 PDF ==========
def extract_data_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

        if not text:
            print("❌ 无法提取文本，可能是扫描版 PDF。")
            return None, None, None

        print("📜 PDF 解析的文本内容：\n", text)

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


# ========== 2. 读取价格表 ==========
def load_price_list(price_path):
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


# ========== 3. 填充 INVOICE ==========
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


# ========== 4. 填充 PACKING LIST ==========
def fill_packing_list(template_path, output_path, part_numbers, ordered_quantities, po_number, price_list):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    ws["K11"] = po_number  # 填充 PO 号

    start_row = 17
    for i, (part_number, ordered_quantity) in enumerate(zip(part_numbers, ordered_quantities)):
        row = start_row + i
        price, units_per_case, nw, gw = price_list.get(part_number, ("N/A", 250, "N/A", "N/A"))

        # **确保 Part Number 填充到 A, B, C 合并单元格**
        merged_cell_range = f"A{row}:C{row}"
        ws.merge_cells(merged_cell_range)  # **合并 A, B, C**
        ws[f"A{row}"] = part_number  # **填充 Part Number**

        ws[f"D{row}"] = f"=F{row} * P{row}"  # **Excel 公式 `=F17 * P17`**
        ws[f"F{row}"] = ordered_quantity
        ws[f"H{row}"] = f"=F{row} * N{row}"  # **Excel 公式 `=F17 * N17`（价格计算）**
        ws[f"N{row}"] = price  # **N 列提取 Clark 表格 E 列数据（单价）**
        ws[f"O{row}"] = nw  # **O 列提取 Clark 表格 D 列数据（净重 NW）**
        ws[f"P{row}"] = units_per_case

    wb.save(output_path)
    print(f"✅ PACKING LIST 生成成功：{output_path}")


# ========== 5. 主程序 ==========
if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "/tmp/temp_po.pdf"
    price_list_path = sys.argv[2] if len(sys.argv) > 2 else "/tmp/temp_price.xlsx"

    invoice_template = "/tmp/INVOICE.xlsx"
    packing_list_template = "/tmp/PACKING_LIST.xlsx"

    invoice_output = "/tmp/INVOICE_2024-00-90868.xlsx"
    packing_list_output = "/tmp/PACKING_LIST_2024-00-90868.xlsx"

    part_numbers, ordered_quantities, po_number = extract_data_from_pdf(pdf_path)

    if part_numbers and ordered_quantities and po_number:
        price_list = load_price_list(price_list_path)
        fill_invoice(invoice_template, invoice_output, part_numbers, ordered_quantities, po_number, price_list)
        fill_packing_list(packing_list_template, packing_list_output, part_numbers, ordered_quantities, po_number, price_list)

        print("🎉 所有文件生成完毕！")
