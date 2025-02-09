import streamlit as st
import os
import shutil

# ========== 1. 设置 Streamlit 页面 ==========
st.title("📄 PO 订单处理工具")
st.write("上传 PO PDF 和 价格表 Excel，自动生成 INVOICE 和 PACKING LIST。")

# ========== 2. 上传 `PO PDF` 文件 ==========
po_file = st.file_uploader("📂 上传 PO 文件（PDF）", type=["pdf"])
if po_file:
    po_path = "PO2024-00-90868.pdf"
    with open(po_path, "wb") as f:
        f.write(po_file.getbuffer())
    st.success("✅ PO 文件已上传！")

# ========== 3. 上传 `价格表 Excel` 文件 ==========
price_file = st.file_uploader("📂 上传价格表文件（Excel）", type=["xlsx"])
if price_file:
    price_path = "Clark11款纸袋报价更新.xlsx"
    with open(price_path, "wb") as f:
        f.write(price_file.getbuffer())
    st.success("✅ 价格表文件已上传！")

# ========== 4. 生成 INVOICE & PACKING LIST ==========
if st.button("🚀 生成 INVOICE 和 PACKING LIST"):
    if not os.path.exists("PO2024-00-90868.pdf") or not os.path.exists("Clark11款纸袋报价更新.xlsx"):
        st.error("❌ 请先上传 PO 文件 和 价格表！")
    else:
        os.system("python3 template_filler.py")  # 运行 `template_filler.py`
        
        # 检查 INVOICE.xlsx 是否生成成功
        if os.path.exists("INVOICE_2024-00-90868.xlsx"):
            st.success("✅ INVOICE 生成成功！")
            st.download_button("⬇️ 下载 INVOICE", open("INVOICE_2024-00-90868.xlsx", "rb"), "INVOICE.xlsx")
        else:
            st.error("❌ INVOICE 生成失败，请检查 `template_filler.py` 是否正确运行！")
        
        # 检查 PACKING LIST.xlsx 是否生成成功
        if os.path.exists("PACKING_LIST_2024-00-90868.xlsx"):
            st.success("✅ PACKING LIST 生成成功！")
            st.download_button("⬇️ 下载 PACKING LIST", open("PACKING_LIST_2024-00-90868.xlsx", "rb"), "PACKING_LIST.xlsx")
        else:
            st.error("❌ PACKING LIST 生成失败，请检查 `template_filler.py` 是否正确运行！")
