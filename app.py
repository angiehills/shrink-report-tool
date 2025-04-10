import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

st.title("📊 Shrink Report PDF to Excel Converter")
st.write("Upload a shrink report PDF and download a clean Excel spreadsheet.")

uploaded_file = st.file_uploader("📄 Choose a PDF file", type="pdf")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    all_data = {}

    # Define consistent column headers
    columns = [
        "Conf #", "Date", "User", "UPC", "Description", "Size", "Reason", "Vendor",
        "Price", "Weight", "Units/Scans", "Retail", "Total"
    ]

    # Loop through all pages except the summary (we’ll detect that dynamically)
    for page_num, page in enumerate(doc, 1):
        text = page.get_text("text")
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        # Skip summary page (we'll handle it later)
        if any(re.search(r"Department\s+Reason\s+Items\s+Total", line, re.IGNORECASE) for line in lines):
            continue

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
            keywords = ["DELI", "BAKERY", "PRODUCE", "MEAT", "GROCERY", "DAIRY", "HBC"]
            found = next((k for k in keywords if any(k in line.upper() for line in lines)), None)
            department = found if found else f"PAGE_{page_num}"

        # Identify where the data starts
        try:
            header_index = next(i for i, line in enumerate(lines) if re.search(r"Conf #|Date|User", line, re.IGNORECASE))
            data_lines = lines[header_index + 1:]
        except StopIteration:
            continue

        # Stop parsing at “Total”
        try:
            total_index = next(i for i, line in enumerate(data_lines) if line.lower().startswith("total"))
            data_lines = data_lines[:total_index]
        except StopIteration:
            pass

        # Group lines into record blocks
        record_blocks = []
        for line in data_lines:
            if re.match(r"\d{5,}-\d{2}", line) or re.match(r"\d{5,}$", line):
                record_blocks.append([line])
            elif record_blocks:
                record_blocks[-1].append(line)

        # Parse fields from blocks
        parsed_rows = []
        for block in record_blocks:
            combined = " ".join(block)
            fields = re.split(r"\s{2,}", combined)
            parsed_rows.append((fields + [""] * len(columns))[:len(columns)])

        df = pd.DataFrame(parsed_rows, columns=columns)

        # Create metadata section
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
        all_data[department] = full_df

    # Dynamically detect and process the summary page
    summary_page_index = None
    for i, page in enumerate(doc):
        page_text = page.get_text("text")
        if re.search(r"Department\s+Reason\s+Items\s+Total", page_text, re.IGNORECASE):
            summary_page_index = i
            break

    if summary_page_index is not None:
        summary_lines = [line.strip() for line in doc[summary_page_index].get_text("text").split("\n") if line.strip()]
        summary_data = []
        capture = False
        for line in summary_lines:
            if re.search(r"Department\s+Reason\s+Items\s+Total", line):
                capture = True
                summary_data.append(["Department", "Reason", "Items", "Total Retail", "Total Cost"])
            elif "Total:" in line:
                total_match = re.findall(r"[\d,]+\.\d{2}", line)
                if total_match:
                    summary_data.append(["", "", "", total_match[0], ""])
                break
            elif capture:
                parts = re.split(r"\s{2,}", line)
                if len(parts) >= 4:
                    summary_data.append(parts[:5])
                else:
                    summary_data.append(parts + [""] * (5 - len(parts)))

        summary_meta = pd.DataFrame([["Shrink Report Summary"]])
        summary_df = pd.DataFrame(summary_data)
        final_summary_sheet = pd.concat([summary_meta, pd.DataFrame([[]]), summary_df], ignore_index=True)
        all_data["Summary"] = final_summary_sheet

    # Create Excel download
    pdf_name = uploaded_file.name.replace(".pdf", "").replace(".PDF", "")
    excel_name = f"{pdf_name}.xlsx"
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in all_data.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False)

    st.success("✅ Conversion complete!")
    st.download_button(
        label="📥 Download Excel File",
        data=output.getvalue(),
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
