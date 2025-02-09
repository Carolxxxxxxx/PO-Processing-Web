import streamlit as st
import os
import subprocess

# 📌 设置文件存储路径
PO_PATH = "PO_LATEST.pdf"
PRICE_PATH = "Clark11款纸袋报价更新.xlsx"

# 🎯 **Streamlit Web 页面**
st.title("📄 PO 订单处理工具")
st.write("📂 **上传 PO PDF 和 价格表 Excel，自动生成 INVOICE 和 PACKING LIST。**")

# 📌 **上传 PO 文件**
po_uploaded = False
po_file = st.file_uploader("📂 上传 PO 文件（PDF）", type=["pdf"])
if po_file:
    with open(PO_PATH, "wb") as f:
        f.write(po_file.getbuffer())
    po_uploaded = True
    st.success("✅ PO 文件已上传！")

# 📌 **上传价格表**
price_uploaded = False
price_file = st.file_uploader("📂 上传价格表文件（Excel）", type=["xlsx"])
if price_file:
    with open(PRICE_PATH, "wb") as f:
        f.write(price_file.getbuffer())
    price_uploaded = True
    st.success("✅ 价格表文件已上传！")

# 📌 **处理文件**
if st.button("🚀 生成 INVOICE 和 PACKING LIST"):
    if not po_uploaded or not price_uploaded:
        st.error("❌ 请先上传 PO 文件 和 价格表！")
    else:
        st.info("⏳ 处理中，请稍等...")
        
        # **调用 template_filler.py**
        result = subprocess.run(["python3", "template_filler.py", PO_PATH, PRICE_PATH], capture_output=True, text=True)

        # **检查 `template_filler.py` 的输出**
        st.text_area("📜 运行日志", result.stdout + result.stderr, height=200)

        if result.returncode != 0:
            st.error("❌ `template_filler.py` 运行失败，请检查日志！")
        else:
            # **检查并提供下载链接**
            invoice_file = "INVOICE_LATEST.xlsx"
            packing_list_file = "PACKING_LIST_LATEST.xlsx"

            if os.path.exists(invoice_file):
                st.success("✅ INVOICE 生成成功！")
                with open(invoice_file, "rb") as f:
                    st.download_button("⬇️ 下载 INVOICE", f, file_name="INVOICE.xlsx")
            else:
                st.error("❌ INVOICE 生成失败，请检查 `template_filler.py` 是否正确运行！")

            if os.path.exists(packing_list_file):
                st.success("✅ PACKING LIST 生成成功！")
                with open(packing_list_file, "rb") as f:
                    st.download_button("⬇️ 下载 PACKING LIST", f, file_name="PACKING_LIST.xlsx")
            else:
                st.error("❌ PACKING LIST 生成失败，请检查 `template_filler.py` 是否正确运行！")

# 📌 **显示开发者信息**
st.write("👨‍💻 **开发者:** Carol")
st.write("🔗 **GitHub:** [PO-Processing-Web](https://github.com/Carolxxxxxxx/PO-Processing-Web)")
