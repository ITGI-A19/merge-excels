import streamlit as st
import pandas as pd
import zipfile
import os
from io import BytesIO
import tempfile

st.title("📦 Excel Merger from ZIP with Header Row Selection")

uploaded_zip = st.file_uploader("Upload a ZIP file containing Excel files", type=["zip"])

if uploaded_zip:
    with tempfile.TemporaryDirectory() as tmpdirname:
        zip_path = os.path.join(tmpdirname, "uploaded.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdirname)

        # Detect Excel files
        excel_files = []
        for root, _, files in os.walk(tmpdirname):
            for file in files:
                if file.lower().endswith((".xlsx", ".xls", ".xlsm")):
                    excel_files.append(os.path.join(root, file))

        if not excel_files:
            st.error("No Excel files found in the uploaded ZIP.")
        else:
            # Preview first file + first sheet to help user choose header row
            preview_file = excel_files[0]
            xls_preview = pd.ExcelFile(preview_file)
            sheet_preview = xls_preview.sheet_names[0]
            df_preview = xls_preview.parse(sheet_preview, header=None)

            st.write("### Preview of First Sheet to Choose Header Row")
            st.dataframe(df_preview.head(10))  # Show top 10 rows

            header_row_index = st.number_input(
                "Select the header row number (1-based index)",
                min_value=1,
                max_value=10,
                value=1,
                step=1
            ) - 1  # Convert to 0-based index

            if st.button("🔄 Merge All Files"):
                combined_df = pd.DataFrame()

                for file in excel_files:
                    try:
                        xls = pd.ExcelFile(file)
                        for sheet in xls.sheet_names:
                            df = xls.parse(sheet, header=header_row_index)
                            combined_df = pd.concat([combined_df, df], ignore_index=True)
                    except Exception as e:
                        st.warning(f"Skipped {os.path.basename(file)} due to error: {e}")

                if not combined_df.empty:
                    st.success("Merged data from all sheets successfully!")
                    st.dataframe(combined_df.head())

                    def convert_df(df):
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Master')
                        return output.getvalue()

                    st.download_button(
                        label="📥 Download Merged Excel File",
                        data=convert_df(combined_df),
                        file_name="merged_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No data merged. Please check header row selection or file content.")
