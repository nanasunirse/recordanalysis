# st.set_page_config(page_title="メインページ", page_icon='icon.png') #show icon
# st.title("Multiple OSS Access Log Analyzer")

import streamlit as st
uploaded_file = st.file_uploader("upload your food log")

import streamlit as st
import tempfile
from pathlib import Path

from final_mynetdiary_html_detail_report import analyze_one_file

st.set_page_config(
    page_title="Nutrition Report Generator",
    layout="wide"
)

st.title("FoodTrack Nutrition Report Generator")

uploaded_file = st.file_uploader(
    "Upload MyNetDiary Excel File",
    type=["xlsx", "xlsm"]
)

if uploaded_file is not None:

    st.success(f"Uploaded: {uploaded_file.name}")

    if st.button("Generate Report"):

        with st.spinner("Analyzing..."):

            # save uploaded excel temporarily
            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".xlsm"
            ) as tmp:

                tmp.write(uploaded_file.getvalue())
                temp_excel_path = Path(tmp.name)

            # output folder
            output_dir = Path("reports")
            output_dir.mkdir(exist_ok=True)

            # run your existing function
            html_report = analyze_one_file(
                temp_excel_path,
                output_dir
            )

        st.success("Report generated successfully!")

        # show html report
        with open(html_report, "r", encoding="utf-8") as f:
            html_content = f.read()

        st.components.v1.html(
            html_content,
            height=1200,
            scrolling=True
        )

        # download button
        with open(html_report, "rb") as f:

            st.download_button(
                label="Download HTML Report",
                data=f,
                file_name=html_report.name,
                mime="text/html"
            )
