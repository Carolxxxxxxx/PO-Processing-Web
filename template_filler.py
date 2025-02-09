import pdfplumber
import openpyxl
import re
import os

# ========== 1. 解析 PDF，提取 PO 号、Part Number、箱数 ==========
def extract_data_from_pdf(pdf_path):
    """
    解析 PDF 提取 Part Number、订单数量 和 PO 号
    """
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

        if not text:
            print("❌ 无法提取文本，可能是扫描版 PDF。")
            return None, None, None

        print("📜 PDF 解析的文本内容：\n", text)

        part_numbers = []
        ordered_quantities = []
        po_number = None

        # 1️⃣ 提取 PO 号
        po_match = re.search(r"Order Number\s*([\d-]+)", text)
        if po_match:
            po_number = po_match.group(1)

        # 2️⃣ 解析 Part Number 和 订单数量
        lines = text.split("\n")

        for i in range(len(lines)):
            # 仅匹配 `BHBxxxx-CLRK` 或 `BHWxxxx-CLRK`
            match = re.search(r"(BHB\d{3,}-CLRK|BHW\d{3,}-CLRK)", lines[i])
            if match:
                part_number = match.group(1)

                # 🔹 确保订单数量（箱数）是数字
                ordered_match = re.findall(r"(\d{2,}\.00)", lines[i])  # 只匹配类似 `600.00`
                if ordered_match:
                    ordered_quantity = int(float(ordered_match[-1]))  # 获取最后一个匹配数值
                else:
                    ordered_quantity = "解析错误"
                    print(f"⚠️ 无法解析数量：{lines[i]}")

                # **确保 `Part Number` 和 `订单数量` 绑定**
                part_numbers.append(part_number)
                ordered_quantities.append(ordered_quantity)

        return part_numbers, ordered_quantities, po_number

# ========== 2. 读取价格表 ==========
def load_price_list(price_path):
    """
    读取价格表，构建 {Part Number: (Price, Units per Case, NW, GW)} 字典
    """
    wb = openpyxl.load_workbook(price_path)
    ws = wb.active
    price_dict = {}

    for row in ws.iter_rows(min_row=2, values_only=True):  # 跳过标题行
        if row[0]:
            part_number = row[0]
            price = row[2] if isinstance(row[2], (int, float)) else "N/A"
            units_per_case = 250 if "Case-250" in row[1] else 200  # 判断包装单位
            nw = row[3] if isinstance(row[3], (int, float)) else "N/A"  # 净重（NW）
            gw = row[4] if isinstance(row[4], (int, float)) else "N/A"  # 毛重（GW）
            price_dict[part_number] = (price, units_per_case, nw, gw)

    return price_dict

# ========== 3. 填充 INVOICE ==========
def fill_invoice(template_path, output_path, part_numbers, ordered_quantities, po_number, price_list):
    """
    填充 INVOICE.xlsx：
    - B15, B16, B17... 填入 Part Number
    - C15, C16, C17... 填入 订单数量（箱数）
    - E15, E16, E17... 填入 产品数量（箱数 × 单箱装量）
    - H15, H16, H17... 填入 价格（匹配价格表）
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    ws["J9"] = po_number  # 填充 PO 号

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
    """
    填充 PACKING LIST.xlsx：
    - K11 填入 PO 号
    - A17, A18...（合并单元格）填入 Part Number
    - D17, D18... **填充 Excel 公式 `=F17 * P17`**
    - F17, F18... 填入 箱数
    - H17, H18... **填充 Excel 公式 `=F17 * N17`（价格计算）**
    - N17, N18... **从 `Clark11款纸袋报价更新.xlsx` 提取 E 列数据（单价）**
    - O17, O18... **从 `Clark11款纸袋报价更新.xlsx` 提取 D 列数据（净重 NW）**
    - P17, P18... 填入 一箱装多少只
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    ws["K11"] = po_number  # 填充 PO 号

    start_row = 17
    for i, (part_number, ordered_quantity) in enumerate(zip(part_numbers, ordered_quantities)):
        row = start_row + i
        price, units_per_case, nw, gw = price_list.get(part_number, ("N/A", 250, "N/A", "N/A"))

        ws[f"A{row}"] = part_number
        ws[f"D{row}"] = f"=F{row} * P{row}"
        ws[f"F{row}"] = ordered_quantity
        ws[f"H{row}"] = f"=F{row} * N{row}"
        ws[f"N{row}"] = price
        ws[f"O{row}"] = nw
        ws[f"P{row}"] = units_per_case

    wb.save(output_path)
    print(f"✅ PACKING LIST 生成成功：{output_path}")

# ========== 5. 主程序 ==========
if __name__ == "__main__":
    pdf_path = "PO2024-00-90868.pdf"
    price_list_path = "Clark11款纸袋报价更新.xlsx"
    invoice_template = "INVOICE.xlsx"
    packing_list_template = "PACKING_LIST.xlsx"

    part_numbers, ordered_quantities, po_number = extract_data_from_pdf(pdf_path)

    if part_numbers and ordered_quantities and po_number:
        price_list = load_price_list(price_list_path)
        fill_invoice(invoice_template, f"INVOICE_{po_number}.xlsx", part_numbers, ordered_quantities, po_number, price_list)
        fill_packing_list(packing_list_template, f"PACKING_LIST_{po_number}.xlsx", part_numbers, ordered_quantities, po_number, price_list)

        print("🎉 所有文件生成完毕！")
