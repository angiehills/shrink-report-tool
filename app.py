import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import os
from io import BytesIO

st.title("\U0001F4CA Shrink Report PDF to Excel Converter")
st.write("Upload a shrink report PDF and download a clean Excel spreadsheet.")

uploaded_file = st.file_uploader("\U0001F4C4 Choose a PDF file", type="pdf")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    all_data = {}

    for page_num, page in enumerate(doc, 1):
        lines = [line.strip() for line in page.get_text().split('\n') if line.strip()]

        # Extract metadata
        store_line = next((line for line in lines if "Piggly" in line), "")
        report_line = next((line for line in lines if "Report" in line), "")
        page_line = next((line for line in lines if "Page #" in line or "Page:" in line), f"Page {page_num}")
        timestamp_line = next((line for line in lines if "/" in line and ":" in line), "")
        department_line = next((line for line in lines if "Department:" in line), None)

        department = ""
        if department_line:
            department = department_line.split(":")[-1].strip().upper()
        else:
            keywords = ["DELI", "BAKERY", "PRODUCE", "MEAT"]
            found = next((k for k in keywords if any(k in line.upper() for line in lines)), None)
            department = found if found else f"PAGE_{page_num}"

        # Skip non-shrink pages
        if "End Reports" in report_line:
            continue

        # Look for structured rows after the column headers
        try:
            header_idx = next(i for i, l in enumerate(lines) if l.lower().startswith("conf #")) + 1
        except StopIteration:
            header_idx = 0

        # Try to split each line into fields using flexible spacing
        data_rows = []
        for line in lines[header_idx:]:
            if line.lower().startswith("total"):
                break
            if len(line.split()) > 4:  # basic validation
                row = re.split(r"\s{2,}|\t", line)
                data_rows.append(row)

        if not data_rows:
            continue

        # Define expected column headers
        columns = [
            "Conf #", "Date", "User", "UPC", "Description", "Size", "Reason", "Vendor",
            "Price Adj", "Weight", "Units/Scans", "Retail/Avg", "Total", "Reclaim Eligible", "Allow Credit"
        ]

        df = pd.DataFrame(data_rows)
        df.columns = columns[:df.shape[1]]

        # Prepend metadata and structure sheet
        metadata = [
            ["Grocery Order Tracking"],
            ["Shrink"],
            [f"Store: {store_line}"],
            [f"Page: {page_line}"],
            [f"Report: {report_line}"],
            [f"Date Printed: {timestamp_line}"],
            [f"Department: {department}"],
            []
        ]
        meta_df = pd.DataFrame(metadata)
        columns_row = pd.DataFrame([df.columns.tolist()])
        total_row = pd.DataFrame([["Total"] + ["" for _ in range(df.shape[1] - 1)]], columns=df.columns)

        full_df = pd.concat([meta_df, columns_row, df, total_row], ignore_index=True)
        all_data[department or f"Page_{page_num}"] = full_df

    if all_data:
        pdf_name = uploaded_file.name.replace(".pdf", "").replace(".PDF", "")
        excel_name = f"{pdf_name}.xlsx"
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet, df in all_data.items():
                df.to_excel(writer, sheet_name=sheet[:31], index=False, header=False)

        st.success("\u2705 Conversion complete!")
        st.download_button(
            label="\U0001F4E5 Download Excel File",
            data=output.getvalue(),
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("\u26A0\uFE0F No valid shrink data found in this PDF. Please check the file format.")
