import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import io

st.set_page_config(page_title="Shrink Report Extractor", layout="centered")
st.title("Shrink Report PDF to Excel Converter")
st.write("Upload a shrink report PDF and download a clean Excel spreadsheet.")

uploaded_file = st.file_uploader("📄 Choose a PDF file", type="pdf")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    lines = []
    for page in doc:
        lines += page.get_text().split("\n")

    lines = [line.strip() for line in lines if line.strip()]
    records = []
    record = {}
    fields_captured = 0

    for i, line in enumerate(lines):
        if re.match(r'^\d{11,}$', line):
            if record and fields_captured >= 7:
                records.append(record)
            record = {}
            fields_captured = 0
            record["UPC"] = line
            fields_captured += 1
        elif "AWG" in line or "NASH" in line:
            record["Vendor"] = line
            fields_captured += 1
        elif re.match(r'^\d+(\.\d{2})$', line):
            if "Price" not in record:
                record["Price"] = line
            else:
                record["Retail"] = line
            fields_captured += 1
        elif re.match(r'^\d+$', line):
            record["Units"] = line
            fields_captured += 1
        elif re.match(r'^\d{2}/\d{2}$', line):
            record["Date"] = line
            fields_captured += 1
        elif re.match(r'^[A-Z]{3}$', line):
            record["User"] = line
            fields_captured += 1
        elif "Out of Date" in line or "Spoilage" in line:
            record["Reason"] = line
            fields_captured += 1
        elif line.lower() not in ["department", "total:"]:
            if "Description" not in record:
                record["Description"] = line
                fields_captured += 1

    if record and fields_captured >= 7:
        records.append(record)

    df = pd.DataFrame(records)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    if not df.empty:
        st.success("✅ Data extracted successfully!")
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="shrink_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No matching shrink data found in the PDF.")
