# excel_manager.py - Handles Excel file management

import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook


def list_excels(folder):
    """Lists all Excel files in the given folder."""
    return [f for f in os.listdir(folder) if f.endswith(('.xls', '.xlsx'))]


def load_excel(file_path):
    """Loads the Excel file and returns the workbook object."""
    return load_workbook(file_path)


def save_excel(workbook, file_path):
    """Saves the modified workbook without deleting existing sheets."""
    workbook.save(file_path)
    st.success("File saved successfully!")


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

        workbook = load_excel(file_path)
        sheet_names = workbook.sheetnames
        selected_sheet = st.selectbox("Select a Sheet", sheet_names)
        sheet = workbook[selected_sheet]

        # Convert sheet to DataFrame
        data = sheet.values
        columns = next(data)  # Extract column names
        df = pd.DataFrame(data, columns=columns)

        # Editable Excel Table
        edited_df = st.data_editor(df, num_rows="dynamic")

        # Sheet Operations
        st.subheader("Sheet Operations")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Add Sheet"):
                new_sheet_name = st.text_input("Enter new sheet name:")
                if new_sheet_name:
                    workbook.create_sheet(title=new_sheet_name)
                    save_excel(workbook, file_path)
        with col2:
            if st.button("Delete Sheet"):
                if len(workbook.sheetnames) > 1:
                    workbook.remove(workbook[selected_sheet])
                    save_excel(workbook, file_path)
                else:
                    st.warning("Cannot delete the last remaining sheet!")
        with col3:
            if st.button("Rename Sheet"):
                new_name = st.text_input("Enter new sheet name:")
                if new_name:
                    sheet.title = new_name
                    save_excel(workbook, file_path)

        # Column Operations
        st.subheader("Column Operations")
        col4, col5 = st.columns(2)
        with col4:
            if st.button("Add Column"):
                col_name = st.text_input("Enter column name:")
                if col_name:
                    edited_df[col_name] = ""
        with col5:
            if st.button("Delete Column"):
                col_to_delete = st.selectbox("Select column to delete", edited_df.columns)
                edited_df.drop(columns=[col_to_delete], inplace=True)

        # Row Operations
        st.subheader("Row Operations")
        col6, col7 = st.columns(2)
        with col6:
            if st.button("Add Row"):
                edited_df.loc[len(edited_df)] = ["" for _ in range(len(edited_df.columns))]
        with col7:
            if st.button("Delete Row"):
                row_to_delete = st.number_input("Enter row index to delete:", min_value=0, max_value=len(edited_df) - 1,
                                                step=1)
                edited_df.drop(index=row_to_delete, inplace=True)

        # Save Changes
        st.subheader("Save Changes")
        if st.button("Save Changes"):
            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
                for cell in row:
                    cell.value = None

            for col_idx, col_name in enumerate(edited_df.columns, start=1):
                sheet.cell(row=1, column=col_idx, value=col_name)

            for row_idx, row in enumerate(edited_df.itertuples(index=False), start=2):
                for col_idx, value in enumerate(row, start=1):
                    sheet.cell(row=row_idx, column=col_idx, value=value)

            save_excel(workbook, file_path)
    else:
        st.warning("No Excel files found in the folder.")
