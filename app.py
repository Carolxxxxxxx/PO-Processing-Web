import streamlit as st
import os
import pandas as pd

# ========== 1. 确保 `INVOICE.xlsx` 存在 ==========
def check_invoice_file():
    """
    检查 `INVOICE.xlsx` 是否存在，如果不存在，则创建一个默认模板
    """
    if not os.path.exists("INVOICE.xlsx"):
        wb = pd.ExcelWriter("INVOICE.xlsx", engine="openpyxl")
        wb.close()
        st.warning("⚠️ `INVOICE.xlsx` 文件未找到，已自动创建一个空白模板！请上传您的 `INVOICE.xlsx` 文件或重新运行 `template_filler.py`。")
    else:
        st.success("✅ `INVOICE.xlsx` 文件已找到，准备就绪！")


# ========== 2. Streamlit 页面 ==========
st.title("📄 PO 订单处理工具")
st.write("使用此工具自动解析 PO 文件，并生成 INVOICE 和 PACKING LIST。")

# ========== 3. 检查 `INVOICE.xlsx` 是否存在 ==========
check_invoice_file()

# ========== 4. 允许用户上传 `INVOICE.xlsx` ==========
uploaded_file = st.file_uploader("📂 上传 `INVOICE.xlsx` 文件", type=["xlsx"])
if uploaded_file:
    with open("INVOICE.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success("✅ `INVOICE.xlsx` 文件已成功上传！请刷新页面并重新运行程序。")

# ========== 5. 运行 `template_filler.py` 生成 INVOICE ==========
if st.button("🚀 生成 INVOICE 和 PACKING LIST"):
    os.system("python3 template_filler.py")
    if os.path.exists("INVOICE_2024-00-90868.xlsx"):
        st.success("✅ INVOICE 生成成功！")
        st.download_button("⬇️ 下载 INVOICE", open("INVOICE_2024-00-90868.xlsx", "rb"), "INVOICE.xlsx")
    else:
        st.error("❌ 生成 INVOICE 失败，请检查 `template_filler.py` 是否正确运行！")

if os.path.exists("PACKING_LIST_2024-00-90868.xlsx"):
    st.success("✅ PACKING LIST 生成成功！")
    st.download_button("⬇️ 下载 PACKING LIST", open("PACKING_LIST_2024-00-90868.xlsx", "rb"), "PACKING_LIST.xlsx")

# ========== 6. 提示用户上传 `INVOICE.xlsx` ==========
st.info("📌 如果遇到 `INVOICE.xlsx` 丢失的问题，请上传文件或检查 `template_filler.py` 是否正确运行！")
