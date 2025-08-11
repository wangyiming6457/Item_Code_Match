import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz import process, fuzz
from io import BytesIO

# --- AUTH ---
def login():
    st.title("🔒 Cost Sheet Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username == "ademco" and password == "yimingiscool":
            st.session_state.logged_in = True
        else:
            st.error("Invalid username or password.")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    login()
    st.stop()

# --- Utilities ---
def clean_model(text):
    if not isinstance(text, str):
        return ""
    return (
        text.strip()
        .lower()
        .replace(" ", "")
        .replace("-", "")
        .replace("\n", "")
        .split("(")[0]
    )

def find_header_row_and_model_col(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=40):
        for cell in row:
            if isinstance(cell.value, str) and "model no" in cell.value.lower():
                return cell.row, cell.column
    return None, None

# --- Streamlit UI ---
st.title("⚡ Cost Sheet Item Code Matcher")
st.write("Upload your Item Listing and Cost Sheet Excel files.")

item_file = st.file_uploader("Upload Item Listing Excel", type="xlsx")
cost_file = st.file_uploader("Upload Cost Sheet Excel (with multiple sheets)", type="xlsx")

if item_file and cost_file:
    with st.spinner("Processing... Please wait..."):
        # Load item listing
        item_listing_df = pd.read_excel(item_file)
        item_lookup = {
            clean_model(name): (name, code)
            for name, code in zip(item_listing_df['Display Name'], item_listing_df['Name'])
            if isinstance(name, str) and clean_model(name) != ""
        }

        # Precompute keys for fuzzy matching
        item_keys = list(item_lookup.keys())

        # Load cost workbook
        original_wb = load_workbook(cost_file)
        output_wb = Workbook()
        output_wb.remove(output_wb.active)

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        progress_bar = st.progress(0)
        total_sheets = len(original_wb.sheetnames)

        for idx, sheet_name in enumerate(original_wb.sheetnames, start=1):
            sheet = original_wb[sheet_name]
            header_row, model_col = find_header_row_and_model_col(sheet)
            if not header_row or not model_col:
                continue

            # Convert sheet to DataFrame
            data = sheet.values
            cols = next(data)
            df = pd.DataFrame(data, columns=cols)

            # Insert "Item Code" column
            insert_at = model_col - 1
            df.insert(insert_at, "Item Code", "")

            # Matching logic
            for i, raw_model in df.iloc[:, insert_at + 1].items():
                if not isinstance(raw_model, str) or not raw_model.strip():
                    continue

                cleaned_model = clean_model(raw_model)
                if cleaned_model == "" or cleaned_model.isdigit():
                    continue

                if cleaned_model in item_lookup:
                    df.iat[i, insert_at] = item_lookup[cleaned_model][1]
                else:
                    best_match = process.extractOne(cleaned_model, item_keys, scorer=fuzz.token_set_ratio)
                    if best_match and best_match[1] >= 90:
                        df.iat[i, insert_at] = item_lookup[best_match[0]][1]
                    else:
                        df.iat[i, insert_at] = None

            # Save DataFrame back to Excel sheet (no heavy style copying)
            new_sheet = output_wb.create_sheet(title=sheet_name)
            for r_idx, row in enumerate([df.columns.tolist()] + df.values.tolist(), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = new_sheet.cell(row=r_idx, column=c_idx, value=value)
                    if r_idx > 1 and c_idx == insert_at + 1 and value in (None, ""):
                        cell.fill = yellow_fill

            progress_bar.progress(idx / total_sheets)

        # Prepare download
        output = BytesIO()
        output_wb.save(output)
        st.success("✅ Done! Download your processed file below.")
        st.download_button(
            label="📥 Download Updated Cost Sheet",
            data=output.getvalue(),
            file_name="CostSheet_ItemCode_Matched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


