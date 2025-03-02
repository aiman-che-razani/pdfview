import streamlit as st
import os
import base64

def display_pdf(source_folder, selected_pdf):
    st.subheader("📕 PDF Viewer")
    pdf_path = os.path.join(source_folder, selected_pdf)

    with open(pdf_path, "rb") as pdf_file:
        base64_pdf = base64.b64encode(pdf_file.read()).decode("utf-8")

    pdf_display = f"""
        <iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700px"></iframe>
    """
    st.markdown(pdf_display, unsafe_allow_html=True)
