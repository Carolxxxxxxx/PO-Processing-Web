import streamlit as st
import os

# ========== 设置 Streamlit 页面 ==========
st.title("📄 PO 订单处理工具")
st.write("上传 PO PDF 和 价格表 Excel，自动生成 INVOICE 和 PACKING LIST。")

# **全局变量：存储最新上传的文件**
PO_PATH = "PO_LATEST.pdf"  
PRICE_PATH = "Clark11款纸袋报价更新.xlsx"

# ========== **1. 上传 `PO PDF` 文件** ==========
po_uploaded = False  # **新增变量，标记 PO 是否上传**
po_file = st.file_uploader("📂 上传 PO 文件（PDF）", type=["pdf"])
if po_file:
    with open(PO_PATH, "wb") as f:
        f.write(po_file.getbuffer())  # **保存新上传的 PO**
    po_uploaded = True
    st.success("✅ PO 文件已上传！")

# ========== **2. 上传 `价格表 Excel` 文件** ==========
price_uploaded = False  # **新增变量，标记价格表是否上传**
price_file = st.file_uploader("📂 上传价格表文件（Excel）", type=["xlsx"])
if price_file:
    with open(PRICE_PATH, "wb") as f:
        f.write(price_file.getbuffer())  # **保存新上传的价格表**
    price_uploaded = True
    st.success("✅ 价格表文件已上传！")

# ========== **3. 运行 `template_filler.py` 生成 INVOICE ==========
if st.button("🚀 生成 INVOICE 和 PACKING LIST"):
    if not po_uploaded or not price_uploaded:
        st.error("❌ 请先上传 PO 文件 和 价格表！")
    else:
        st.info("⏳ 正在解析 PO 并生成文件，请稍等...")
        result = os.system(f"python3 template_filler.py {PO_PATH} {PRICE_PATH}")  # **确保读取最新 PO**
        
        if result != 0:
            st.error("❌ `template_filler.py` 运行失败，请检查日志！")
        else:
            # **检查 INVOICE 是否生成**
            if os.path.exists("INVOICE_最新.xlsx"):
                st.success("✅ INVOICE 生成成功！")
                st.download_button("⬇️ 下载 INVOICE", open("INVOICE_最新.xlsx", "rb"), "INVOICE.xlsx")
            else:
                st.error("❌ INVOICE 生成失败，请检查 `template_filler.py` 是否正确运行！")

            # **检查 PACKING LIST 是否生成**
            if os.path.exists("PACKING_LIST_最新.xlsx"):
                st.success("✅ PACKING LIST 生成成功！")
                st.download_button("⬇️ 下载 PACKING LIST", open("PACKING_LIST_最新.xlsx", "rb"), "PACKING_LIST.xlsx")
            else:
                st.error("❌ PACKING LIST 生成失败，请检查 `template_filler.py` 是否正确运行！")
