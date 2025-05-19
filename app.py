import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# Page setup
st.set_page_config(page_title="Invoice Table Extractor", layout="wide")
st.title("PDF Invoice to Excel Generator")

# File uploader
uploaded_file = st.file_uploader("Upload your Invoice PDF", type="pdf")

# Function to extract and combine tables from all pages
def extract_all_tables(file):
    combined_table = pd.DataFrame()
    log = []

    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    if table and len(table) > 1:
                        try:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            combined_table = pd.concat([combined_table, df], ignore_index=True)
                            log.append(f"[✅ Success] Page {i+1}: {len(df)} rows extracted.")
                        except Exception as e:
                            log.append(f"[❌ Error] Page {i+1}: Failed to process table - {e}")
                    else:
                        log.append(f"[⚠️ Warning] Page {i+1}: Table found, but it was empty.")
            else:
                log.append(f"[ℹ️ Info] Page {i+1}: No table found.")
    
    return combined_table, log

# Function to save combined table to Excel
def save_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Combined_Table', index=False)
    output.seek(0)
    return output

# Main logic
if uploaded_file:
    with st.spinner("🔍 Extracting tables from all pages..."):
        combined_df, logs = extract_all_tables(uploaded_file)

        st.subheader("📋 Extraction Log")
        for log in logs:
            st.text(log)

        if not combined_df.empty:
            st.success(f"✅ Extracted a total of {len(combined_df)} rows from the PDF.")
            st.write("📌 Preview of Combined Table:")
            st.dataframe(combined_df)

            excel_data = save_to_excel(combined_df)

            st.download_button(
                label="📥 Download as Excel",
                data=excel_data,
                file_name="combined_invoice_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No tables were found or extracted from the uploaded PDF.")
