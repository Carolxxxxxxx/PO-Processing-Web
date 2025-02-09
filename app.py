import streamlit as st
import os
import time

st.title("📜 PO 处理网站")
st.write("请上传 PO 文件和价格表，系统会自动生成 INVOICE 和 PACKING LIST")

# 上传 PO PDF 文件
uploaded_pdf = st.file_uploader("📂 上传 PO 文件", type=["pdf"])
# 上传价格表
uploaded_price_list = st.file_uploader("📂 上传价格表", type=["xlsx"])

if uploaded_pdf and uploaded_price_list:
    # **使用 `/tmp/` 目录存储临时文件（适用于 Streamlit Cloud）**
    pdf_path = "/tmp/temp_po.pdf"
    price_path = "/tmp/temp_price.xlsx"

    with open(pdf_path, "wb") as f:
        f.write(uploaded_pdf.getbuffer())

    with open(price_path, "wb") as f:
        f.write(uploaded_price_list.getbuffer())

    st.success("✅ 文件上传成功，点击按钮开始处理！")

    if st.button("🚀 处理 PO 并生成 Excel"):
        # **确保 `template_filler.py` 在当前目录**
        if os.path.exists("template_filler.py"):
            # **执行 `template_filler.py` 并等待处理完成**
            result = os.system(f"python3 template_filler.py {pdf_path} {price_path}")
            time.sleep(3)

            # **检查是否生成了文件**
            invoice_path = "/tmp/INVOICE_2024-00-90868.xlsx"
            packing_list_path = "/tmp/PACKING_LIST_2024-00-90868.xlsx"

            if os.path.exists(invoice_path):
                with open(invoice_path, "rb") as f:
                    st.download_button("📥 下载 INVOICE.xlsx", f, file_name="INVOICE.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("❌ `INVOICE.xlsx` 文件未找到，请检查 `template_filler.py` 是否正常运行！")

            if os.path.exists(packing_list_path):
                with open(packing_list_path, "rb") as f:
                    st.download_button("📥 下载 PACKING_LIST.xlsx", f, file_name="PACKING_LIST.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("❌ `PACKING_LIST.xlsx` 文件未找到，请检查 `template_filler.py` 是否正常运行！")
        else:
            st.error("❌ `template_filler.py` 文件不存在，请确保它在当前目录！")
