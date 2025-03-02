import streamlit as st
from sidebar import load_sidebar
from excel_manager import manage_excel
from pdf_viewer import display_pdf

st.set_page_config(layout="wide")

st.title("📁 Folder-Based PDF & Excel Viewer & Editor")

# Sidebar
source_folder, selected_pdf, selected_excel = load_sidebar()

# PDF Viewer
if selected_pdf:
    display_pdf(source_folder, selected_pdf)

# Excel Viewer & Manager
if selected_excel:
    manage_excel(source_folder, selected_excel)
