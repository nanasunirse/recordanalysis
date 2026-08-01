import streamlit as st
import tempfile
from pathlib import Path

from final_mynetdiary_html_detail_report import analyze_one_file

st.title("Nutrition Report Generator")

uploaded_file = st.file_uploader(
    "Upload MyNetDiary File",
    type=["xlsx", "xlsm"]
)

if uploaded_file:

    if st.button("Generate Report"):

        # save uploaded file
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsm"
        ) as tmp:

            tmp.write(uploaded_file.getvalue())
            excel_path = Path(tmp.name)

        output_dir = Path("output")

        # run your existing python code
        report_path = analyze_one_file(
            excel_path,
            output_dir
        )

        st.success("Report Generated")

        # read html
        html_content = Path(report_path).read_text(
            encoding="utf-8"
        )

        # display html
        st.components.v1.html(
            html_content,
            height=1500,
            scrolling=True
        )

import streamlit as st
uploaded_file = st.file_uploader("upload your food log")
