# st.set_page_config(page_title="メインページ", page_icon='icon.png') #show icon
# st.title("Multiple OSS Access Log Analyzer")

import streamlit as st
uploaded_file = st.file_uploader("upload your food log")

if uploaded_file is not None:
wb = load_workbook(
uploaded_file,
data_only=True,
keep_vba=True
)
st.success("File loaded successfully")
