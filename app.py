import streamlit as st
import pandas as pd
import tempfile
import os
import Lot_extraction as le
import Additional_Word_file_processing as afp
# import csv_to_excel as cte
import Tables_to_CSV as ttc
import CSV_to_excel as cte
import subprocess, shutil
import traceback

# resaving doc
def normalize_docx(upload_path):
    libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
    
    # on linux (deployed), use 'libreoffice'; on windows use full path
    import platform
    if platform.system() == "Windows":
        lo = libreoffice_path
    else:
        lo = "libreoffice"
    
    try:
        out_dir = tempfile.mkdtemp()
        subprocess.run([lo, "--headless", "--convert-to", "docx",
                        "--outdir", out_dir, upload_path], check=True)
        converted = [f for f in os.listdir(out_dir) if f.endswith(".docx")][0]
        return os.path.join(out_dir, converted)
    except (FileNotFoundError, subprocess.CalledProcessError):
        # LibreOffice not available, return original path
        return upload_path
    
# page config
st.set_page_config(page_title="Tender File Processing", page_icon="📄", layout="centered")
st.title("📄 Tender File Processing")
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None
    st.session_state.intermediate_tables = None
st.markdown('<p style="font-size:20px;">Upload your Excel and Word files below. The app will extract the needed information and combine it into one Excel file.</p>', unsafe_allow_html=True)

# ---- File Uploaders ----
st.header("Step 1: Upload Your Files")

excel_file = st.file_uploader(
    "Upload Excel files (.xlsx, .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=False
)

main_word_file = st.file_uploader(
    "Upload Main Word file (.docx)",
    type=["docx"],
    accept_multiple_files=False
)

additional_word_file = st.file_uploader(
    "Upload Additional Word File (.docx)",
    type=["docx"],
    accept_multiple_files=False
)
# ---- Process Button ----
st.header("Step 2: Process Files")

if st.button("Extract & Combine Data", type="primary"):

    # Make sure the user actually uploaded something
    if not main_word_file:
        st.warning("Please upload at least one file before processing.")
    else:
        with st.spinner("Processing your files..."):
            try:
        # Save uploads to temp files
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp.write(main_word_file.read())
                    main_word_path = normalize_docx(tmp.name)

            # st.write(f"DEBUG path: {main_word_path}")
            # st.write(f"DEBUG size: {os.path.getsize(main_word_path)} bytes")

                tables = le.save_relevant_tables(main_word_path, "relevant_tables")
            except Exception as e:
                st.error(f"❌ Failed to process main Word file: {e}")
                st.code(traceback.format_exc())
                st.stop()
                # st.write(f"DEBUG save_relevant_tables returned {len(tables)} tables")
            # from docx import Document
            # _doc = Document(main_word_path)
            # st.write(f"DEBUG tables in doc: {len(_doc.tables)}")
            # st.write(f"DEBUG paragraphs in doc: {len(_doc.paragraphs)}")
            try: 
                if excel_file:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        tmp.write(excel_file.read())
                        excel_path = tmp.name
                    excel_table = ttc.load_mixed_tables([excel_path])
                    tables.extend(excel_table)
            except Exception as e:
                st.error(f"❌ Failed to process Excel file: {e}")
                st.code(traceback.format_exc())
                st.stop()
            try:
                if additional_word_file:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                        tmp.write(additional_word_file.read())
                        additional_word_path = normalize_docx(tmp.name)
                    additional_word_table = afp.save_csv(additional_word_path)
                    tables.extend(additional_word_table)
            except Exception as e:
                st.error(f"❌ Failed to process additional Word file: {e}")
                st.code(traceback.format_exc())
                st.stop()
            #extracted_data = ttc.extract_data(tables)
            # st.write("tables from le.save_relevant_tables")
            # for i, t in enumerate(tables[:len(tables)]):
            #     st.write(f"Item{i}: type={type(t).__name__}")
            #     if isinstance(t, pd.DataFrame):
            #         st.write(f" shape={t.shape}, columns={list(t.columns)}")
            #         st.dataframe(t.head(3))
            #     else:
            #         st.write(f"value preview: {str(t)[:200]}")
            st.session_state.tables = tables

            try:
                extracted_data = ttc.extract_data(tables)
                st.session_state.extracted_data = extracted_data
                st.session_state.intermediate_tables = tables
            except Exception as e:
                st.error(f"❌ Failed to extract data: {e}")
                st.code(traceback.format_exc())
                st.stop()

        st.success("Done, preview below.")

# preview
if st.session_state.extracted_data is not None:
    st.header("Step 3: Preview")

    data = st.session_state.extracted_data

    if isinstance(data, pd.DataFrame):
        df = data

    else:
        df = pd.DataFrame(data)

    st.write(f"**Rows:** {len(df)} **Columns:** {len(df.columns)}")
    st.dataframe(df, use_container_width=True)

    # download as excel

    st.header("Step 4: Download")

    # sheet_name = st.text_input("Sheet name", value="Extracted Data")
    file_name = st.text_input("File name", value="extracted_data")
    output = cte.format_excel(df)

    st.download_button(
        label="Download as Excel",
        data=output,
        file_name=f"{file_name if file_name else 'extracted_data'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )