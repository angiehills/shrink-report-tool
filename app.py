import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO

st.title("📊 Shrink Report PDF to Excel Converter")
st.write("Upload a shrink report PDF and download a clean Excel spreadsheet.")

uploaded_file = st.file_uploader("📄 Choose a PDF file", type="pdf")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")

    # Define headers and refined column mapping
    columns_final = [
        "Conf #", "Date", "User", "UPC", "Description", "Size", "Reason", "Vendor",
        "Price", "Weight", "Units/Scans", "Retail/Avg", "Total"
    ]

    def final_map_x_to_column(x):
        if x < 50: return "Conf #"
        elif x < 80: return "Date"
        elif x < 120: return "User"
        elif x < 220: return "UPC"
        elif x < 400: return "Description"
        elif x < 440: return "Size"
        elif x < 500: return "Reason"
        elif x < 560: return "Vendor"
        elif x < 600: return "Price"
        elif x < 640: return "Weight"
        elif x < 680: return "Units/Scans"
        elif x < 720: return "Retail/Avg"
        else: return "Total"

    pages_data = {}
    pages = list(doc.pages())

    for page in pages:
        blocks = page.get_text("dict")["blocks"]
        row_data = defaultdict(dict)

        store = report = date = dept = page_label = ""

        for b in blocks:
            for l in b.get("lines", []):
                y = round(l["bbox"][1], 1)
                for s in l["spans"]:
                    x = s["bbox"][0]
                    text = s["text"].strip()
                    if not text:
                        continue
                    col = final_map_x_to_column(x)
                    row_data[y][col] = text

                    # Pull metadata on the fly
                    if "Piggly" in text:
                        store = text
                    elif "Report" in text:
                        report = text
                    elif "/" in text and ":" in text:
                        date = text
                    elif "Department" in text:
                        dept = text.split(":")[-1].strip()
                    elif "Page" in text:
                        page_label = text

        # Fallbacks
        if not dept:
            dept = f"Page_{page.number+1}"

        # Extract only clean rows with a valid Conf #
        clean_structured_rows = []
        for y in sorted(row_data.keys()):
            row = row_data[y]
            if re.match(r"\d{5,}-\d{2}", row.get("Conf #", "")):
                clean_structured_rows.append([row.get(c, "") for c in columns_final])

        # Build full DataFrame with metadata + headers + data
        df = pd.DataFrame(clean_structured_rows, columns=columns_final)
        meta = pd.DataFrame([
            ["Grocery Order Tracking"],
            ["Shrink"],
            [f"Store: {store}"],
            [f"Page: {page_label}"],
            [f"Report: {report}"],
            [f"Date Printed: {date}"],
            [f"Department: {dept}"],
            []
        ])
        header = pd.DataFrame([columns_final])
        total_row = pd.DataFrame([["Total"] + [""] * (len(columns_final) - 1)], columns=columns_final)

        full_tab = pd.concat([meta, header, df, total_row], ignore_index=True)
        pages_data[dept] = full_tab

    # Export to Excel
    pdf_name = uploaded_file.name.replace(".pdf", "").replace(".PDF", "")
    excel_name = f"{pdf_name}_converted.xlsx"
    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for dept, df in pages_data.items():
            df.to_excel(writer, sheet_name=dept[:31], index=False, header=False)

    st.success("✅ Conversion complete!")
    st.download_button(
        label="📥 Download Excel File",
        data=output.getvalue(),
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
