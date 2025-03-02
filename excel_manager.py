# excel_manager.py - Handles Excel file management

import streamlit as st
import pandas as pd
import os

def list_excels(folder):
    """Lists all Excel files in the given folder."""
    return [f for f in os.listdir(folder) if f.endswith(('.xls', '.xlsx', '.csv'))]

def excel_manager_ui(folder):
    """Streamlit UI for managing Excel files."""
    st.header("Excel Manager")

    if not os.path.exists(folder):
        st.error(f"Folder '{folder}' not found.")
        return

    excel_files = list_excels(folder)

    if excel_files:
        selected_excel = st.selectbox("Select an Excel file", excel_files)
        file_path = os.path.join(folder, selected_excel)

        try:
            df = pd.read_excel(file_path) if selected_excel.endswith(('.xls', '.xlsx')) else pd.read_csv(file_path)
            st.dataframe(df)

            if st.button("Download File"):
                with open(file_path, "rb") as f:
                    st.download_button("Download", f, file_name=selected_excel)
        except Exception as e:
            st.error(f"Error reading file: {e}")
    else:
        st.warning("No Excel files found in the folder.")
