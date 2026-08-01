# st.set_page_config(page_title="メインページ", page_icon='icon.png') #show icon
# st.title("Multiple OSS Access Log Analyzer")

import streamlit as st
uploaded_file = st.file_uploader("upload your food log")

import pandas as pd
if uploaded_file is not None:
    df = pd.read_csv(
        uploaded_file,
        sep=r'\s(?=(?:[^"]*"[^"]*")*[^"]*$)(?![^\[]*\])',
        engine='python',
        na_values='-',
        header=None)
