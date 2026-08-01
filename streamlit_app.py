# st.set_page_config(page_title="メインページ", page_icon='icon.png') #show icon
# st.title("Multiple OSS Access Log Analyzer")

import streamlit as st
import pandas as pd

uploaded_file = st.file_uploader(
    "Upload your food log",
    type=["xlsx", "xlsm"]
)

import openpyxl
import streamlit as st

st.write("openpyxl installed!")
st.write(openpyxl.__version__)

import streamlit as st

try:
    import openpyxl
    st.success("openpyxl installed")
except Exception as e:
    st.error(e)




if uploaded_file is not None:

    df = pd.read_excel(
        uploaded_file,
        header=None
    )

    st.write(df.head(20))


