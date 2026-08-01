# st.set_page_config(page_title="メインページ", page_icon='icon.png') #show icon
# st.title("Multiple OSS Access Log Analyzer")


import streamlit as st
from pathlib import Path

uploaded_file = st.file_uploader(
    "Upload your food log",
    type=["xlsx", "xlsm"]
)

if uploaded_file is not None:

    temp_path = Path("uploaded_report.xlsm")

    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.write(temp_path)

    report = parse_workbook(temp_path)
