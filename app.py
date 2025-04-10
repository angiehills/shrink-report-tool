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
    parsed_departments = defaultdict(list)

    # Extract blocks with coordinates to reconstruct lines
    for page_num, page in enumerate(doc, 1):
        blocks = page.get_text("blocks")
        text_blocks = [(round(b[1], 1), b[4]) for b in blocks if b[4].strip()]
        grouped_rows = defaultdict(list)
        for y, text in text_blocks:
            grouped_rows[y].append(text)

        # Try to find department name
        department = f"Page_{page_num}"
        for _, texts in grouped_rows.items():
            if any("Department" in t for t in texts):
                for t in texts:
                    if "Department" not in t:
                        department = t.strip().upper()
                        break

        # Skip summary page for now
        if any("Reason" in " ".join(v) and "Items" in " ".join(v) for v in grouped_rows.values()):
            continue

        # Add lines containing shrink data
        for y in sorted(grouped_rows.keys()):
            line = " ".join(grouped_rows[y])
            if re.search(r"\d{5,}-\d{2}", line):  # Conf #
                parsed_departments[department].append(line)

    # Define final column order from the PDF
    columns = [
        "Conf #", "Date", "User", "UPC", "Description", "Size", "Reason", "Vendor",
        "Price", "Weight", "Units/Scans", "Retail/Avg", "Total"
    ]

    structured_data = {}
    for dept, lines in parsed_departments.items():
        parsed_rows = []
        for raw in lines:
            parts = raw.replace("\n", " ").split()
            if len(parts) < 14:
                continue
            try:
                conf_idx = next(i for i, p in enumerate(parts) if re.match(r"\d{5,}-\d{2}", p))
                conf = parts[conf_idx]
                date = parts[conf_idx + 1]
                user = parts[conf_idx + 2]
                upc_idx = next(i for i, p in enumerate(parts) if re.match(r"\d{11,}", p))
                upc = parts[upc_idx]
                size = parts[upc_idx - 1]
                description = " ".join(parts[conf_idx + 3:upc_idx - 1])
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
                    conf, date, user, upc, description, size, reason, vendor,
                    price, weight, units, retail, total
                ]
                parsed_rows.append((row + [""] * len(columns))[:len(columns)])
            except Exception:
                continue

        df = pd.DataFrame(parsed_rows, columns=columns)
        structured_data[dept] = df

    if structured_data:
        pdf_name = uploaded_file.name.replace(".pdf", "").replace(".PDF", "")
        excel_name = f"{pdf_name}_parsed.xlsx"
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet, df in structured_data.items():
                df.to_excel(writer, sheet_name=sheet[:31], index=False)

        st.success("✅ Conversion complete!")
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid shrink data found in this PDF.")
