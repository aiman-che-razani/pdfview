import streamlit as st
import os
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
import io
import base64
import tempfile

# Set page layout
st.set_page_config(layout="wide")

st.title("📁 Folder-Based PDF & Excel Viewer Web App")

# ---- LEFT SIDEBAR: FOLDER NAVIGATION ----
st.sidebar.header("📂 Select Source Folder")

# User inputs a folder path
source_folder = st.sidebar.text_input("Enter Folder Path:", "")

if source_folder and os.path.exists(source_folder):
    # Get first-level subfolders
    first_level_folders = [f for f in os.listdir(source_folder) if os.path.isdir(os.path.join(source_folder, f))]
    
    if first_level_folders:
        selected_first_level = st.sidebar.selectbox("🔹 Select First-Level Folder:", first_level_folders)

        # Get second-level subfolders inside the selected first-level folder
        first_level_path = os.path.join(source_folder, selected_first_level)
        second_level_folders = [f for f in os.listdir(first_level_path) if os.path.isdir(os.path.join(first_level_path, f))]
        
        if second_level_folders:
            selected_second_level = st.sidebar.selectbox("🔸 Select Second-Level Folder:", second_level_folders)
            full_folder_path = os.path.join(first_level_path, selected_second_level)
        else:
            full_folder_path = first_level_path  # If no second-level folder, use first-level folder
    else:
        full_folder_path = source_folder  # If no first-level folder, use main source folder

    # Get all files in the selected folder
    files = os.listdir(full_folder_path)
    pdf_files = [f for f in files if f.endswith(".pdf")]
    excel_files = [f for f in files if f.endswith((".xlsx", ".xls"))]

    selected_pdf = st.sidebar.selectbox("📄 Select a PDF File:", pdf_files) if pdf_files else None
    selected_excel = st.sidebar.selectbox("📊 Select an Excel File:", excel_files) if excel_files else None

else:
    st.sidebar.warning("⚠ Please enter a valid folder path.")

# ---- MAIN DISPLAY AREA ----
col1, col2 = st.columns([11, 9])

# ---- PDF Viewer ----
with col1:
    st.subheader("📕 PDF Viewer")

    if selected_pdf:
        pdf_path = os.path.join(full_folder_path, selected_pdf)

        # Convert PDF to base64 for embedding
        with open(pdf_path, "rb") as pdf_file:
            base64_pdf = base64.b64encode(pdf_file.read()).decode("utf-8")

        pdf_display = f"""
            <iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700px"></iframe>
        """
        st.markdown(pdf_display, unsafe_allow_html=True)
    else:
        st.info("📂 Select a PDF from the sidebar.")

# ---- Excel Viewer (Editable) ----
with col2:
    st.subheader("📊 Excel Viewer & Editor")

    if selected_excel:
        excel_path = os.path.join(full_folder_path, selected_excel)

        # Load Excel file
        excel_data = pd.ExcelFile(excel_path)
        sheet_name = st.selectbox("Select Sheet", excel_data.sheet_names)

        if sheet_name:
            df = pd.read_excel(excel_data, sheet_name=sheet_name)

            # Configure AgGrid (Editable Table)
            gb = GridOptionsBuilder.from_dataframe(df)
            gb.configure_default_column(editable=True)  # Make all columns editable
            grid_options = gb.build()

            # Display editable table
            grid_response = AgGrid(df, gridOptions=grid_options, height=400, fit_columns_on_grid_load=True)
            updated_df = grid_response['data']  # Get edited data

            # Button to download the updated file
            if st.button("Save Changes & Download"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    updated_df.to_excel(writer, sheet_name=sheet_name, index=False)
                output.seek(0)

                st.download_button(label="📥 Download Updated Excel", 
                                   data=output, 
                                   file_name="updated_excel.xlsx", 
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("📂 Select an Excel file from the sidebar.")
