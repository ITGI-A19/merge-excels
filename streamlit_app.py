import streamlit as st
import pandas as pd
import zipfile
import os
from io import BytesIO
import tempfile
import time

st.set_page_config(page_title="Excel ZIP Merger", layout="centered")
st.title("📦 Excel Merger from ZIP with Header Selection & Progress")

uploaded_zip = st.file_uploader("Upload a ZIP file containing Excel files", type=["zip"])

if uploaded_zip:
    st.info("Unzipping and scanning Excel files...")
    try:
        with tempfile.TemporaryDirectory() as tmpdirname:
            zip_path = os.path.join(tmpdirname, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())

            # Extract ZIP
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdirname)

            # Find Excel files
            excel_files = []
            for root, _, files in os.walk(tmpdirname):
                for file in files:
                    if file.lower().endswith((".xlsx", ".xls", ".xlsm")):
                        excel_files.append(os.path.join(root, file))

            if not excel_files:
                st.error("No Excel files found in the ZIP.")
            else:
                # Preview first file/sheet
                preview_file = excel_files[0]
                xls_preview = pd.ExcelFile(preview_file)
                df_preview = xls_preview.parse(xls_preview.sheet_names[0], header=None)
                st.write("### Preview First Sheet")
                st.dataframe(df_preview.head(10))

                header_row_index = st.number_input(
                    "Select the header row number (1-based)", min_value=1, max_value=10, value=1, step=1
                ) - 1

                if st.button("🔄 Start Merging"):
                    combined_df = pd.DataFrame()
                    file_progress = st.progress(0)
                    status = st.empty()

                    total_files = len(excel_files)
                    errors = []
                    skip_count = 0

                    for i, file_path in enumerate(excel_files):
                        try:
                            xls = pd.ExcelFile(file_path)
                            for sheet in xls.sheet_names:
                                df = xls.parse(sheet, header=header_row_index)
                                if not df.empty:
                                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                        except Exception as e:
                            errors.append(f"{os.path.basename(file_path)}: {e}")
                            skip_count += 1
                        if i % 2 == 0 or i == total_files - 1:
                            file_progress.progress((i + 1) / total_files)
                            status.text(f"Processed {i + 1}/{total_files} files...")

                    if not combined_df.empty:
                        st.success(f"✅ Merged {total_files - skip_count} files. Skipped {skip_count}.")
                        st.dataframe(combined_df.head())

                        with st.spinner("📦 Creating downloadable Excel..."):
                            def to_excel(df):
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    df.to_excel(writer, index=False, sheet_name='Master')
                                return output.getvalue()

                            final_excel = to_excel(combined_df)

                        st.download_button(
                            label="📥 Download Merged Excel File",
                            data=final_excel,
                            file_name="merged_output.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                        if errors:
                            st.expander("⚠️ Skipped Files (click to view)").write("\n".join(errors))
                    else:
                        st.error("Merging failed. No data found or all sheets were empty.")
    except Exception as e:
        st.error(f"❌ Unexpected error: {e}")
        st.stop()

    st.caption("⚠️ If your ZIP is over 200MB, consider running this app locally for better performance.")
