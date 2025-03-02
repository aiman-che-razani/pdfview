# app.py - Main entry point

import streamlit as st
import os
from pdf_viewer import pdf_viewer_ui
from excel_manager import excel_manager_ui


def get_second_level_folders(root_folder):
    """Retrieves all first-layer and second-layer subfolders inside the root folder."""
    if not os.path.exists(root_folder):
        return [], {}

    first_level_folders = [d for d in os.listdir(root_folder) if os.path.isdir(os.path.join(root_folder, d))]
    second_level_folders = {}

    for folder in first_level_folders:
        second_level_folders[folder] = [
            d for d in os.listdir(os.path.join(root_folder, folder))
            if os.path.isdir(os.path.join(root_folder, folder, d))
        ]

    return first_level_folders, second_level_folders


def main():
    st.set_page_config(page_title='PDF & Excel Manager', layout='wide')

    st.sidebar.title("Settings")
    root_folder = st.sidebar.text_input("Root Folder Path", "./source_files")

    first_level_folders, second_level_folders = get_second_level_folders(root_folder)

    if not first_level_folders:
        st.sidebar.warning("No first-layer folders found in the root directory.")
        return

    first_layer_selected = st.sidebar.selectbox("Select a First-Layer Folder", first_level_folders)

    if first_layer_selected and first_layer_selected in second_level_folders:
        second_layer_folders = second_level_folders[first_layer_selected]
        second_layer_selected = st.sidebar.selectbox("Select a Second-Layer Folder", second_layer_folders) if second_layer_folders else None
    else:
        second_layer_selected = None

    st.title("PDF & Excel Manager")
    st.write(f"**Root Path:** {root_folder}")
    st.write(f"**First-Layer Folder:** {first_layer_selected if first_layer_selected else 'None'}")
    st.write(f"**Second-Layer Folder:** {second_layer_selected if second_layer_selected else 'None'}")

    if second_layer_selected:
        selected_folder = os.path.join(root_folder, first_layer_selected, second_layer_selected)
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
