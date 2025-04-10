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
    departments_full = {}
    summary_sheet = None

    def extract_metadata(lines):
        store = next((l for l in lines if "Piggly" in l), "")
        report = next((l for l in lines if "Report" in l), "")
        page = next((l for l in lines if "Page" in l), "")
        date = next((l for l in lines if "/" in l and ":" in l), "")
        dept_line = next((l for l in lines if "Department" in l), "")
        dept = dept_line.split(":")[-1].strip() if dept_line else ""
        return store, report, page, date, dept

    for page_num, page in enumerate(doc, 1):
        blocks = page.get_text("blocks")
        lines_raw = [b[4].strip() for b in blocks if b[4].strip()]
        store, report, page_label, date_str, dept_name = extract_metadata(lines_raw)
        y_blocks = defaultdict(list)
        for b in blocks:
            y = round(b[1], 1)
            y_blocks[y].append(b[4])
        rows = [" ".join(y_blocks[y]) for y in sorted(y_blocks)]

        if any("Department" in r and "Reason" in r and "Items" in r for r in rows):
            summary_data = [["Department", "Reason", "Items", "Total Retail", "Total Cost"]]
            for r in rows:
                if "Department" in r and "Reason" in r:
                    continue
                if r.lower().startswith("total"):
                    break
                parts = re.split(r"\s{2,}", r)
                summary_data.append(parts + [""] * (5 - len(parts)))
            summary_sheet = pd.DataFrame(summary_data)
            continue

        column_headers = [
            "Conf #", "Date", "User", "UPC", "Description", "Size", "Reason", "Vendor",
            "Price", "Weight", "Units/Scans", "Retail/Avg", "Total"
        ]
        records = []
        for r in rows:
            parts = r.replace("\n", " ").split()
            try:
                conf_idx = next(i for i, p in enumerate(parts) if re.match(r"\d{5,}-\d{2}", p))
                conf = parts[conf_idx]
                date = parts[conf_idx + 1]
                user = parts[conf_idx + 2]
                upc_idx = next(i for i, p in enumerate(parts) if re.match(r"\d{11,}", p))
                upc = parts[upc_idx]
                size = parts[upc_idx - 1]
                desc = " ".join(parts[conf_idx + 3:upc_idx - 1])
                vendor = parts[upc_idx + 1] + " " + parts[upc_idx + 2]
                units = parts[upc_idx + 3]
                reason = parts[upc_idx + 4]
                if not parts[upc_idx + 5].replace('.', '', 1).isdigit():
                    reason += " " + parts[upc_idx + 5]
                    price_idx = upc_idx + 6
                else:
                    price_idx = upc_idx + 5
                price = parts[price_idx]
                retail = parts[price_idx + 1]
                total = parts[price_idx + 2] if len(parts) > price_idx + 2 else ""
                weight = parts[price_idx + 3] if len(parts) > price_idx + 3 and not parts[price_idx + 3].replace('.', '', 1).isdigit() else ""
                row = [
                    conf, date, user, upc, desc, size, reason, vendor,
                    price, weight, units, retail, total
                ]
                records.append((row + [""] * len(column_headers))[:len(column_headers)])
            except Exception:
                continue

        df_data = pd.DataFrame(records, columns=column_headers)
        meta = pd.DataFrame([
            ["Grocery Order Tracking"],
            ["Shrink"],
            [f"Store: {store}"],
            [f"Page: {page_label}"],
            [f"Report: {report}"],
            [f"Date Printed: {date_str}"],
            [f"Department: {dept_name}"],
            []
        ])
        header = pd.DataFrame([column_headers])
        total_row = pd.DataFrame([["Total"] + [""] * (len(column_headers) - 1)], columns=column_headers)
        full_tab = pd.concat([meta, header, df_data, total_row], ignore_index=True)
        tab_name = dept_name if dept_name else f"Page_{page_num}"
        departments_full[tab_name] = full_tab

    # Export final Excel
    pdf_name = uploaded_file.name.replace(".pdf", "").replace(".PDF", "")
    excel_name = f"{pdf_name}_converted.xlsx"
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in departments_full.items():
            df.to_excel(writer, sheet_name=name[:31], index=False, header=False)
        if summary_sheet is not None:
            pd.DataFrame([["Shrink Report Summary"], []]).to_excel(writer, sheet_name="Summary", index=False, header=False)
            summary_sheet.to_excel(writer, sheet_name="Summary", startrow=2, index=False)

    st.success("✅ Conversion complete!")
    st.download_button(
        label="📥 Download Excel File",
        data=output.getvalue(),
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
