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
    data_rows = []

    for page in doc:
        text = page.get_text()
        lines = text.split("\n")
        for line in lines:
            if re.match(r'^\d{5,}\s+.+?\s{2,}.+?\s{2,}.*?\$\d+\.\d{2}\s+\d+\s+\$\d+\.\d{2}$', line):
                parts = re.split(r'\s{2,}', line)
                if len(parts) >= 5:
                    product_info = parts[0].split()
                    upc = product_info[0]
                    product_name = " ".join(product_info[1:])
                    department = parts[1]
                    unit_price = parts[2].replace("$", "")
                    quantity = parts[3]
                    total = parts[4].replace("$", "")
                    data_rows.append([upc, product_name, department, unit_price, quantity, total])

    if data_rows:
        df = pd.DataFrame(data_rows, columns=["UPC", "Product Name", "Department", "Unit Price", "Quantity", "Total"])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.success("✅ Data extracted successfully!")
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="shrink_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No matching shrink data found in the PDF.")
