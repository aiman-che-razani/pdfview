# app.py - Main entry point

import streamlit as st
import os
from pdf_viewer import pdf_viewer_ui
from excel_manager import excel_manager_ui


def get_second_level_folders(root_folder):
    """Retrieves all second-layer subfolders inside the root folder."""
    if not os.path.exists(root_folder):
        return []
    first_level = [os.path.join(root_folder, d) for d in os.listdir(root_folder) if
                   os.path.isdir(os.path.join(root_folder, d))]
    second_level = []
    for folder in first_level:
        second_level.extend(
            [os.path.join(folder, d) for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d))])
    return second_level


def main():
    st.set_page_config(page_title='PDF & Excel Manager', layout='wide')

    st.sidebar.title("Settings")
    root_folder = st.sidebar.text_input("Root Folder Path", "./source_files")

    subfolders = get_second_level_folders(root_folder)
    selected_folder = st.sidebar.selectbox("Select a Subfolder", subfolders) if subfolders else None

    st.title("PDF & Excel Manager")

    if selected_folder:
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("PDF Viewer")
            pdf_viewer_ui(selected_folder)

        with col2:
            st.subheader("Excel Manager")
            excel_manager_ui(selected_folder)
    else:
        st.warning("No valid second-layer subfolders found.")


if __name__ == "__main__":
    main()
