import streamlit as st
import os

st.title("📜 PO 处理网站")
st.write("请上传 PO 文件和价格表，系统会自动生成 INVOICE 和 PACKING LIST")

# 上传 PO PDF 文件
uploaded_pdf = st.file_uploader("📂 上传 PO 文件", type=["pdf"])
# 上传价格表
uploaded_price_list = st.file_uploader("📂 上传价格表", type=["xlsx"])

if uploaded_pdf and uploaded_price_list:
    with open("temp_po.pdf", "wb") as f:
        f.write(uploaded_pdf.getbuffer())

    with open("temp_price.xlsx", "wb") as f:
        f.write(uploaded_price_list.getbuffer())

    st.success("✅ 文件上传成功，点击按钮开始处理！")

    if st.button("🚀 处理 PO 并生成 Excel"):
        os.system("python3 template_filler.py")
        st.success("✅ 处理完成！请下载文件")

        with open("INVOICE_2024-00-90868.xlsx", "rb") as f:
            st.download_button("📥 下载 INVOICE.xlsx", f, file_name="INVOICE.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with open("PACKING_LIST_2024-00-90868.xlsx", "rb") as f:
            st.download_button("📥 下载 PACKING_LIST.xlsx", f, file_name="PACKING_LIST.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
