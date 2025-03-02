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
        columns = next(data, [])  # Extract column names
        df = pd.DataFrame(data, columns=columns) if columns else pd.DataFrame()

        # Editable Excel Table
        edited_df = st.data_editor(df, num_rows="dynamic")

        # Sheet Operations
        st.subheader("Sheet Operations")
        sheet_col1, sheet_col2 = st.columns([2, 1])
        with sheet_col1:
            new_sheet_name = st.text_input("Enter new sheet name:", key="new_sheet_name")
        with sheet_col2:
            rename_sheet_name = st.text_input("Rename sheet:", key="rename_sheet_name")

        sheet_col3, sheet_col4, sheet_col5 = st.columns(3)
        with sheet_col3:
            if st.button("Add Sheet") and new_sheet_name:
                if new_sheet_name not in workbook.sheetnames:
                    workbook.create_sheet(title=new_sheet_name)
                    save_excel(workbook, file_path)
                else:
                    st.warning("Sheet already exists!")
        with sheet_col4:
            if st.button("Delete Sheet"):
                if len(workbook.sheetnames) > 1:
                    workbook.remove(workbook[selected_sheet])
                    save_excel(workbook, file_path)
                else:
                    st.warning("Cannot delete the last remaining sheet!")
        with sheet_col5:
            if st.button("Rename Sheet") and rename_sheet_name:
                sheet.title = rename_sheet_name
                save_excel(workbook, file_path)

        # Column Operations
        st.subheader("Column Operations")
        col4, col5 = st.columns(2)
        with col4:
            new_col_name = st.text_input("Enter column name:", key="new_column_name")
            if st.button("Add Column") and new_col_name:
                edited_df[new_col_name] = ""
        with col5:
            columns_to_delete = st.multiselect("Select columns to delete", edited_df.columns)
            if st.button("Delete Selected Columns") and columns_to_delete:
                edited_df.drop(columns=columns_to_delete, inplace=True)

        # Row Operations
        st.subheader("Row Operations")
        row_col1, row_col2 = st.columns(2)
        with row_col1:
            add_row_button = st.button("Add Row")
        with row_col2:
            row_to_delete = st.number_input("Enter row index to delete:", min_value=0,
                                            max_value=max(len(edited_df) - 1, 0), step=1)
            delete_row_button = st.button("Delete Selected Row")

        if add_row_button:
            edited_df.loc[len(edited_df)] = ["" for _ in range(len(edited_df.columns))]

        if delete_row_button:
            edited_df.drop(index=row_to_delete, inplace=True)

        # Save Changes
        st.subheader("Save Changes")
        if st.button("Save Changes"):
            # Ensure sheet is cleared before writing new data
            sheet.delete_rows(1, sheet.max_row)

            for col_idx, col_name in enumerate(edited_df.columns, start=1):
                sheet.cell(row=1, column=col_idx, value=col_name)

            for row_idx, row in enumerate(edited_df.itertuples(index=False), start=2):
                for col_idx, value in enumerate(row, start=1):
                    sheet.cell(row=row_idx, column=col_idx, value=value)

            save_excel(workbook, file_path)
    else:
        st.warning("No Excel files found in the folder.")
