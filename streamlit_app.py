# st.set_page_config(page_title="メインページ", page_icon='icon.png') #show icon
# st.title("Multiple OSS Access Log Analyzer")

import streamlit as st
import pandas as pd

uploaded_file = st.file_uploader(
    "Upload your food log",
    type=["xlsx", "xlsm"]
)

if uploaded_file is not None:

    df = pd.read_excel(
        uploaded_file,
        header=None
    )

    st.write(df.head(20))
