import streamlit as st
import os

def load_sidebar():
    st.sidebar.header("📂 Select Source Folder")
    source_folder = st.sidebar.text_input("Enter Folder Path:", "")

    selected_pdf, selected_excel = None, None
    if source_folder and os.path.exists(source_folder):
        files = os.listdir(source_folder)
        pdf_files = [f for f in files if f.endswith(".pdf")]
        excel_files = [f for f in files if f.endswith((".xlsx", ".xls"))]

        selected_pdf = st.sidebar.selectbox("📄 Select a PDF File:", pdf_files) if pdf_files else None
        selected_excel = st.sidebar.selectbox("📊 Select an Excel File:", excel_files) if excel_files else None

    return source_folder, selected_pdf, selected_excel
