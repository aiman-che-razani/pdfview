import streamlit as st
import os
import pandas as pd
import io
import base64

# Set page layout
st.set_page_config(layout="wide")

st.title("📁 Folder-Based PDF & Excel Viewer & Editor")

# ---- LEFT SIDEBAR: FOLDER NAVIGATION ----
st.sidebar.header("📂 Select Source Folder")

source_folder = st.sidebar.text_input("Enter Folder Path:", "")

if source_folder and os.path.exists(source_folder):
    first_level_folders = [f for f in os.listdir(source_folder) if os.path.isdir(os.path.join(source_folder, f))]

    if first_level_folders:
        selected_first_level = st.sidebar.selectbox("🔹 Select First-Level Folder:", first_level_folders)

        first_level_path = os.path.join(source_folder, selected_first_level)
        second_level_folders = [f for f in os.listdir(first_level_path) if os.path.isdir(os.path.join(first_level_path, f))]

        if second_level_folders:
            selected_second_level = st.sidebar.selectbox("🔸 Select Second-Level Folder:", second_level_folders)
            full_folder_path = os.path.join(first_level_path, selected_second_level)
        else:
            full_folder_path = first_level_path
    else:
        full_folder_path = source_folder

    files = os.listdir(full_folder_path)
    pdf_files = [f for f in files if f.endswith(".pdf")]
    excel_files = [f for f in files if f.endswith((".xlsx", ".xls"))]

    selected_pdf = st.sidebar.selectbox("📄 Select a PDF File:", pdf_files) if pdf_files else None
    selected_excel = st.sidebar.selectbox("📊 Select an Excel File:", excel_files)

else:
    st.sidebar.warning("⚠ Please enter a valid folder path.")

# ---- TOP SECTION: Excel Management Functions ----
st.subheader("📊 Manage Excel Sheets & Data")

if selected_excel:
    excel_path = os.path.join(full_folder_path, selected_excel)
    excel_data = pd.ExcelFile(excel_path)
    sheet_names = excel_data.sheet_names

    # Select Sheet
    selected_sheet = st.selectbox("Select Sheet", sheet_names, key="sheet_selector")

    # Load DataFrame whenever sheet changes
    if selected_sheet not in st.session_state:
        df = pd.read_excel(excel_data, sheet_name=selected_sheet)
        st.session_state[selected_sheet] = df.copy()
    else:
        df = st.session_state[selected_sheet]

    # ---- Sheet Management: Rename, Add, Delete ----
    col_top1, col_top2, col_top3 = st.columns(3)

    with col_top1:
        new_sheet_name = st.text_input("Rename Sheet:", value=selected_sheet, key="rename_sheet")
        if st.button("✏ Rename Sheet"):
            with pd.ExcelFile(excel_path, engine="openpyxl") as existing_excel:
                sheet_data = {sheet: pd.read_excel(existing_excel, sheet) for sheet in existing_excel.sheet_names}

            sheet_data[new_sheet_name] = sheet_data.pop(selected_sheet)

            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                for sheet, data in sheet_data.items():
                    data.to_excel(writer, sheet_name=sheet, index=False)

            st.success(f"✅ Sheet '{selected_sheet}' renamed to '{new_sheet_name}'")

    with col_top2:
        if st.button("➕ Add Sheet"):
            new_sheet_name = f"Sheet{len(sheet_names) + 1}"
            with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a") as writer:
                pd.DataFrame().to_excel(writer, sheet_name=new_sheet_name, index=False)
            st.success(f"✅ New sheet '{new_sheet_name}' added!")

    with col_top3:
        if len(sheet_names) > 1:
            sheets_to_delete = st.multiselect("Select Sheets to Delete:", sheet_names, key="delete_sheets")
            if st.button("❌ Delete Selected Sheets"):
                with pd.ExcelFile(excel_path, engine="openpyxl") as existing_excel:
                    sheet_data = {sheet: pd.read_excel(existing_excel, sheet) for sheet in existing_excel.sheet_names}

                for sheet in sheets_to_delete:
                    if sheet in sheet_data:
                        sheet_data.pop(sheet)

                with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                    for sheet, data in sheet_data.items():
                        data.to_excel(writer, sheet_name=sheet, index=False)

                st.success(f"✅ Deleted sheets: {', '.join(sheets_to_delete)}")

    # ---- Modify Data Section ----
    st.markdown("### 🛠 Modify Table")
    col_mod1, col_mod2 = st.columns(2)

    with col_mod1:
        if st.button("➕ Add Row"):
            new_row = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
            df = pd.concat([df, new_row], ignore_index=True)
            st.session_state[selected_sheet] = df

        if st.button("➕ Add Column"):
            col_count = len(df.columns) + 1
            while f"Column {col_count}" in df.columns:
                col_count += 1
            new_col_name = f"Column {col_count}"
            df[new_col_name] = ""
            st.session_state[selected_sheet] = df

    with col_mod2:
        if len(df.index) > 0 and st.button("❌ Delete Last Row"):
            df.drop(df.index[-1], inplace=True)
            st.session_state[selected_sheet] = df

        if len(df.columns) > 0:
            col_to_delete = st.selectbox("Select Column to Delete:", df.columns)
            if st.button("❌ Delete Column"):
                df.drop(columns=[col_to_delete], inplace=True)
                st.session_state[selected_sheet] = df

    # ---- Save Changes Without Deleting Other Sheets ----
    st.markdown("### 💾 Save & Download")
    col_top4, col_top5 = st.columns(2)

    with col_top4:
        if st.button("⬇ Push Column Names to First Row & Save"):
            column_headers = df.columns.tolist()
            df.loc[-1] = column_headers
            df.index = df.index + 1
            df = df.sort_index()
            df.columns = [f"Column {i+1}" for i in range(len(column_headers))]
            st.session_state[selected_sheet] = df
            st.success("✅ Column names moved to first row and saved!")

    with col_top5:
        if st.button("💾 Save Changes Without Deleting Other Sheets"):
            with pd.ExcelFile(excel_path, engine="openpyxl") as existing_excel:
                sheet_data = {sheet: pd.read_excel(existing_excel, sheet) for sheet in existing_excel.sheet_names}

            sheet_data[selected_sheet] = df

            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                for sheet, data in sheet_data.items():
                    data.to_excel(writer, sheet_name=sheet, index=False)

            st.success("✅ Changes Saved Without Deleting Other Sheets!")

    # ---- PDF & Excel Viewer ----
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📕 PDF Viewer")
        if selected_pdf:
            pdf_path = os.path.join(full_folder_path, selected_pdf)
            with open(pdf_path, "rb") as pdf_file:
                base64_pdf = base64.b64encode(pdf_file.read()).decode("utf-8")
            st.markdown(f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700px"></iframe>', unsafe_allow_html=True)
        else:
            st.info("📂 Select a PDF from the sidebar.")

    with col2:
        st.subheader("📊 Excel Viewer")
        st.dataframe(st.session_state[selected_sheet])
