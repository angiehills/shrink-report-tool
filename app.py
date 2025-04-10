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
        text = page.get_text("text")
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        # Metadata extraction
        store_line = next((line for line in lines if "Piggly" in line), "")
        report_line = next((line for line in lines if "Report" in line), "")
        page_line = next((line for line in lines if "Page #" in line or "Page:" in line), f"Page {page_num}")
        timestamp_line = next((line for line in lines if "/" in line and ":" in line), "")
        department_line = next((line for line in lines if "Department:" in line), None)

        department = ""
        if department_line:
            department = department_line.split(":")[-1].strip().upper()
        else:
            keywords = ["DELI", "BAKERY", "PRODUCE", "MEAT", "GROCERY"]
            found = next((k for k in keywords if any(k in line.upper() for line in lines)), None)
            department = found if found else f"PAGE_{page_num}"

        if "End Reports" in report_line:
            summary_df = pd.DataFrame([
                ["Grocery Order Tracking"],
                ["Shrink"],
                [f"Store: {store_line}"],
                [f"Page: {page_line}"],
                [f"Report: {report_line}"],
                [f"Date Printed: {timestamp_line}"],
                [f"Department: {department}"],
                [],
                ["Summary or No Shrink Data Found"]
            ])
            all_data[department or f"Page_{page_num}"] = summary_df
            continue

        # Attempt to identify start of data
        header_keywords = ["Conf", "Date", "User"]
        header_idx = next((i for i, line in enumerate(lines) if all(k in line for k in header_keywords)), None)

        data_lines = lines[header_idx + 1:] if header_idx is not None else []

        record_blocks = []
        block = []
        for line in data_lines:
            if line.lower().startswith("total"):
                break
            if re.match(r"\d{5,}-\d{2}", line) or re.match(r"\d{5,}$", line):
                if block:
                    record_blocks.append(block)
                block = [line]
            else:
                block.append(line)
        if block:
            record_blocks.append(block)

        columns = [
            "Conf #", "Date", "User", "UPC", "Description", "Size", "Reason", "Vendor",
            "Price Adj", "Weight", "Units/Scans", "Retail/Avg", "Total"
        ]

        parsed_rows = []
        for block in record_blocks:
            combined = " ".join(block)
            fields = re.split(r"\s{2,}|(?<=\d) (?=\d{5,})", combined.strip())
            row = (fields + [""] * len(columns))[:len(columns)]
            parsed_rows.append(row)

        df = pd.DataFrame(parsed_rows, columns=columns)

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
        columns_row = pd.DataFrame([columns])
        total_row = pd.DataFrame([["Total"] + ["" for _ in range(len(columns) - 1)]], columns=columns)

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
