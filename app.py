import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import os
from io import BytesIO

st.title("📊 Shrink Report PDF to Excel Converter")
st.write("Upload a shrink report PDF and download a clean Excel spreadsheet.")

uploaded_file = st.file_uploader("📄 Choose a PDF file", type="pdf")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    all_data = {}

    for page_num, page in enumerate(doc, 1):
        lines = [line.strip() for line in page.get_text().split('\n') if line.strip()]

        # Extract metadata
        store_line = next((line for line in lines if "Piggly" in line), "")
        report_line = next((line for line in lines if "Report" in line), "")
        page_line = next((line for line in lines if "Page #" in line), f"Page {page_num}")
        timestamp_line = next((line for line in lines if "/" in line and ":" in line), "")
        department_line = next((line for line in lines if "Department:" in line), "Department: Unknown")
        department = department_line.split(":")[-1].strip().upper() or f"Page_{page_num}"

        # Find the index of the header row
        try:
            header_idx = lines.index("Conf #\tDate\tUser\tUPC\tDescription\tSize\tReason\tVendor\tPrice Adj\tWeight\tUnits/Scans\tRetail/Avg Ret Scans\tTotal\tReclaim Eligible\tAllow Credit") + 1
        except:
            try:
                header_idx = next(i for i, l in enumerate(lines) if l.startswith("Conf #")) + 1
            except:
                header_idx = None

        # Extract data rows
        data_rows = []
        for line in lines[header_idx:]:
            if line.lower().startswith("total"):
                break
            if len(line.strip()) > 10:
                row = re.split(r"\s{2,}", line.strip())
                data_rows.append(row)

        df = pd.DataFrame(data_rows)

        # Prepend metadata to top rows
        metadata = [
            ["Grocery Order Tracking"],
            ["Shrink"],
            [f"Store: {store_line}"],
            [f"Page: {page_line}"],
            [f"Report: {report_line}"],
            [f"Date Printed: {timestamp_line}"],
            [f"Department: {department}"],
            [],
        ]
        meta_df = pd.DataFrame(metadata)
        full_df = pd.concat([meta_df, df], ignore_index=True)
        all_data[department] = full_df

    # Output Excel with original name
    pdf_name = uploaded_file.name.replace(".pdf", "").replace(".PDF", "")
    excel_name = f"{pdf_name}.xlsx"
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet, df in all_data.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False, header=False)
    
    st.success("✅ Conversion complete!")
    st.download_button(
        label="📥 Download Excel File",
        data=output.getvalue(),
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
