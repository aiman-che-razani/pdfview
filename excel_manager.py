import streamlit as st
import pandas as pd
import os

def manage_excel(source_folder, selected_excel):
    excel_path = os.path.join(source_folder, selected_excel)
    excel_data = pd.ExcelFile(excel_path)
    sheet_names = excel_data.sheet_names

    selected_sheet = st.selectbox("Select Sheet", sheet_names)

    if selected_sheet not in st.session_state:
        df = pd.read_excel(excel_data, sheet_name=selected_sheet)
        st.session_state[selected_sheet] = df.copy()
    else:
        df = st.session_state[selected_sheet]

    # Modify Data
    col1, col2 = st.columns(2)
    with col1:
        if st.button("➕ Add Row"):
            new_row = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
            df = pd.concat([df, new_row], ignore_index=True)
            st.session_state[selected_sheet] = df

        if st.button("➕ Add Column"):
            col_count = len(df.columns) + 1
            while f"Column {col_count}" in df.columns:
                col_count += 1
            df[f"Column {col_count}"] = ""
            st.session_state[selected_sheet] = df

    with col2:
        if len(df.index) > 0 and st.button("❌ Delete Last Row"):
            df.drop(df.index[-1], inplace=True)
            st.session_state[selected_sheet] = df

        if len(df.columns) > 0:
            col_to_delete = st.selectbox("Select Column to Delete:", df.columns)
            if st.button("❌ Delete Column"):
                df.drop(columns=[col_to_delete], inplace=True)
                st.session_state[selected_sheet] = df

    # Save Excel
    if st.button("💾 Save Changes Without Deleting Other Sheets"):
        with pd.ExcelFile(excel_path, engine="openpyxl") as existing_excel:
            sheet_data = {sheet: pd.read_excel(existing_excel, sheet) for sheet in existing_excel.sheet_names}

        sheet_data[selected_sheet] = df

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for sheet, data in sheet_data.items():
                data.to_excel(writer, sheet_name=sheet, index=False)

        st.success("✅ Changes Saved Without Deleting Other Sheets!")

    # Display DataFrame
    st.subheader("📊 Excel Viewer")
    st.dataframe(st.session_state[selected_sheet])
