# pdf_viewer.py - Handles PDF viewing inside Streamlit

import streamlit as st
import os
import base64
import tempfile


def list_pdfs(folder):
    """Lists all PDF files in the given folder."""
    return [f for f in os.listdir(folder) if f.endswith('.pdf')]


def display_pdf(pdf_path):
    """Displays the selected PDF file inside Streamlit using base64 encoding."""
    with open(pdf_path, "rb") as pdf_file:
        base64_pdf = base64.b64encode(pdf_file.read()).decode("utf-8")

    pdf_display = f"""
        <iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800px"></iframe>
    """
    st.markdown(pdf_display, unsafe_allow_html=True)


def pdf_viewer_ui(folder):
    """Streamlit UI for viewing PDFs."""
    st.header("PDF Viewer")

    if not os.path.exists(folder):
        st.error(f"Folder '{folder}' not found.")
        return

    pdf_files = list_pdfs(folder)

    if pdf_files:
        selected_pdf = st.selectbox("Select a PDF", pdf_files)
        file_path = os.path.join(folder, selected_pdf)
        display_pdf(file_path)
    else:
        st.warning("No PDFs found in the folder.")
