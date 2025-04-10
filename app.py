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
    pages_data = {}
    summary_rows = []

    columns_expected = [
        "Conf #", "Date", "User", "UPC", "Description", "Size", "Reason", "Vendor",
        "Price", "Weight", "Units/Scans", "Retail/Avg", "Total"
    ]

    known_depts = ["BAKERY", "DAIRY", "DELI", "GROCERY", "HBC", "MEAT FRESH", "PRODUCE"]

    for page_num, page in enumerate(doc, 1):
        blocks = page.get_text("blocks")
        store = report = date = dept = page_label = ""
        lines = [b[4].strip() for b in blocks if b[4].strip()]

        # Extract metadata
        for line in lines:
            if "Piggly" in line:
                store = line
            elif "Report" in line:
                report = line
            elif "Page" in line:
                page_label = line
            elif "/" in line and ":" in line:
                date = line

        dept_line = next((b[4] for b in blocks if "Department" in b[4]), "")
        dept = dept_line.split(":")[-1].strip() if dept_line else f"Page_{page_num}"

        # Group by y, sort by x to preserve structure
        rows_by_y = defaultdict(list)
        for b in blocks:
            y = round(b[1], 1)
            rows_by_y[y].append((b[0], b[4].strip()))

        structured_rows = []
        for y in sorted(rows_by_y):
            row = sorted(rows_by_y[y], key=lambda x: x[0])
            text_line = [r[1] for r in row if r[1]]
            structured_rows.append(" ".join(text_line))

        # Detect and parse summary page
        if any("Department" in line and "Items" in line for line in structured_rows):
            for line in structured_rows:
                match = re.match(r"(\d+)\s+(.+?)\s+([\d.]+)\s+([A-Z ]+)", line)
                if match:
                    items, reason, retail, department = match.groups()
                    summary_rows.append([department.strip(), reason.strip(), int(items), float(retail)])
                elif line.lower().startswith("total"):
                    total_val = re.findall(r"[\d,.]+", line)
                    if total_val:
                        summary_rows.append(["", "Total", "", float(total_val[0])])
            continue

        # Parse data rows using fixed field markers
        data_rows = []
        for line in structured_rows:
            parts = line.split()
            if len(parts) >= 13 and any("-" in p for p in parts):
                data_rows.append(parts)

        df_rows = []
        for parts in data_rows:
            try:
                conf_idx = next(i for i, p in enumerate(parts) if re.match(r"\d{5,}-\d{2}", p))
                conf = parts[conf_idx]
                date_field = parts[conf_idx + 1]
                user = parts[conf_idx + 2]
                upc_idx = next(i for i, p in enumerate(parts) if re.match(r"\d{11,}", p))
                upc = parts[upc_idx]
                size = parts[upc_idx - 1]
                description = " ".join(parts[conf_idx + 3:upc_idx - 1])
                vendor = parts[upc_idx + 1] + " " + parts[upc_idx + 2]
                units = parts[upc_idx + 3]
                reason = parts[upc_idx + 4]
                if not parts[upc_idx + 5].replace(".", "", 1).isdigit():
                    reason += " " + parts[upc_idx + 5]
                    price_idx = upc_idx + 6
                else:
                    price_idx = upc_idx + 5
                price = parts[price_idx]
                retail = parts[price_idx + 1]
                total = parts[price_idx + 2] if len(parts) > price_idx + 2 else ""
                weight = parts[price_idx + 3] if len(parts) > price_idx + 3 and not parts[price_idx + 3].replace('.', '', 1).isdigit() else ""
                df_rows.append([
                    conf, date_field, user, upc, description, size, reason, vendor,
                    price, weight, units, retail, total
                ])
            except:
                continue

        df = pd.DataFrame(df_rows, columns=columns_expected)
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
        header = pd.DataFrame([columns_expected])
        total_row = pd.DataFrame([["Total"] + [""] * (len(columns_expected) - 1)], columns=columns_expected)
        full_page = pd.concat([meta, header, df, total_row], ignore_index=True)
        pages_data[dept] = full_page

    # Create clean summary DataFrame
    summary_df = pd.DataFrame(summary_rows, columns=["Department", "Reason", "Items", "Total Retail"])

    # Export to Excel
    pdf_name = uploaded_file.name.replace(".pdf", "").replace(".PDF", "")
    excel_name = f"{pdf_name}_converted.xlsx"
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in pages_data.items():
            df.to_excel(writer, sheet_name=name[:31], index=False, header=False)
        if not summary_df.empty:
            pd.DataFrame([["Shrink Report Summary"], []]).to_excel(writer, sheet_name="Summary", index=False, header=False)
            summary_df.to_excel(writer, sheet_name="Summary", startrow=2, index=False)

    st.success("✅ Conversion complete!")
    st.download_button(
        label="📥 Download Excel File",
        data=output.getvalue(),
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
