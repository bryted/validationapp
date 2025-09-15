import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
import io

st.set_page_config(page_title="MARS KPI Validation", layout="wide")
st.title("📊 MARS Data Quality Validation App")

# --- Step 1: Upload Files ---
st.header("Step 1: Upload Files")
key_file = st.file_uploader("🔑 Upload the KEY file (Excel)", type=["xlsx"])
data_file = st.file_uploader("📂 Upload the DATA file (Excel)", type=["xlsx"])

if key_file and data_file:
    with st.spinner("Reading files..."):
        key_description = pd.read_excel(key_file, sheet_name=None)
        data_sheets = pd.read_excel(data_file, sheet_name=None)

    # --- Step 2: Country & Language Selection ---
    st.header("Step 2: Select Options")
    country_choice = st.selectbox("🌍 Select Country", options=["GHA", "CIV"], index=0)
    language_choice = st.selectbox("🗣️ Select Language", options=["EN", "FR"], index=0)

    # --- Build Answer Map ---
    row_only_fields = []
    row_or_column_fields = ["VisitType", "ReNoSchool", "RChildType", "RHhType", "EducStatus"]

    answer_map = {}
    answer_list_df = key_description.get("answer_list", pd.DataFrame())
    if not answer_list_df.empty:
        answer_list_df.dropna(how="all", inplace=True)
        answer_list_df.dropna(axis=1, how="all", inplace=True)
        grouped = answer_list_df.groupby("Field Name Eng for partner")
        for field, group in grouped:
            if pd.isna(field):
                continue
            all_values = []
            for _, row in group.iterrows():
                if field in row_only_fields:
                    values = row.drop(labels=["Language", "Field Name Eng for partner"]).dropna().tolist()
                elif field in row_or_column_fields:
                    row_vals = row.drop(labels=["Language", "Field Name Eng for partner"]).dropna().tolist()
                    col_vals = group.columns[2:][~row[2:].isna()].tolist()
                    values = row_vals + col_vals
                else:
                    values = group.columns[2:][~row[2:].isna()].tolist()
                all_values.extend([str(v).strip() for v in values if str(v).strip()])
            unique_values = list(dict.fromkeys(all_values))
            if unique_values:
                answer_map[field] = unique_values

    # --- Step 3: Run Validation ---
    if st.button("🚦 Run Validation"):
        with st.spinner("Running validation checks..."):
            description_df = key_description["description"] if "description" in key_description else key_description
            required_meta = description_df[
                description_df["For Mars KPI reporting"].isin(["Y", country_choice])
            ]
            expected_fields = required_meta["Field Name Eng for partner"].dropna().unique().tolist()
            field_types = dict(zip(
                description_df["Field Name Eng for partner"],
                description_df["Type of variable in the table"]
            ))

            data_issues = []

            def log_issue(sheet, field, row, issue_type, description):
                data_issues.append({
                    "Sheet": sheet, "Field": field, "Row": row,
                    "Issue Type": issue_type, "Description": description
                })

            # --- Validation Messages ---
            validation_msgs = {
                "empty_sheet": {"EN": "Sheet is present but contains no data", "FR": "La feuille est présente mais ne contient aucune donnée"},
                "completely_empty": {"EN": "The variable '{}' is completely empty.", "FR": "La variable '{}' est complètement vide."},
                "missing_required": {"EN": "Field is required but missing", "FR": "Champ requis mais manquant"},
                "expected_date": {"EN": "Expected format is DD-MM-YYYY", "FR": "Format attendu : JJ-MM-AAAA"},
                "expected_numeric": {"EN": "Expected a numeric value", "FR": "Une valeur numérique était attendue"},
                "invalid_value": {"EN": "'{}' not in allowed values", "FR": "'{}' ne fait pas partie des valeurs autorisées"},
                "alpha_numeric": {"EN": "Value must be alphanumeric, not digits only", "FR": "La valeur doit être alphanumérique, pas uniquement des chiffres"},
                "duplicate_combo": {"EN": "Duplicate entries found based on combination of: {}", "FR": "Doublons détectés selon la combinaison : {}"},
            }

            def validate_sheet(df, sheet_name):
                if df.empty:
                    log_issue(sheet_name, None, None, "Empty Sheet", validation_msgs["empty_sheet"][language_choice])
                    return
                for field in expected_fields:
                    if field not in df.columns:
                        continue
                    col_data = df[field].replace(r"^\s*$", np.nan, regex=True)
                    if col_data.isnull().all():
                        log_issue(sheet_name, field, None, "Missing Value",
                                  validation_msgs["completely_empty"][language_choice].format(field))
                        continue
                    for idx, val in col_data.items():
                        if pd.isna(val):
                            log_issue(sheet_name, field, idx, "Missing Value",
                                      validation_msgs["missing_required"][language_choice])
                            continue
                        if field in field_types and "date" in str(field_types[field]).lower():
                            try:
                                pd.to_datetime(val, format="%d-%m-%Y")
                            except Exception:
                                log_issue(sheet_name, field, idx, "Date Format Error",
                                          validation_msgs["expected_date"][language_choice])
                        if field in answer_map and str(val).strip().lower() not in [str(v).strip().lower() for v in answer_map[field]]:
                            log_issue(sheet_name, field, idx, "Invalid Value",
                                      validation_msgs["invalid_value"][language_choice].format(val))

            for sheet in ["P", "B-C", "D", "E_Com", "E_Hho", "E_Chd"]:
                if sheet in data_sheets:
                    validate_sheet(data_sheets[sheet], sheet)

        # --- Step 4: Results ---
        st.subheader("Validation Summary")
        if data_issues:
            summary_df = pd.DataFrame(data_issues)
            st.dataframe(summary_df)

            # --- Word Report ---
            word_buffer = io.BytesIO()
            word_doc = Document()
            word_doc.add_heading("MARS Data Quality Report", 0)
            for sheet in summary_df["Sheet"].unique():
                word_doc.add_heading(f"Sheet: {sheet}", level=1)
                issues = summary_df[summary_df["Sheet"] == sheet]
                table = word_doc.add_table(rows=1, cols=4)
                table.style = "Light List Accent 1"
                hdr = table.rows[0].cells
                hdr[0].text = "Field"; hdr[1].text = "Issue Type"; hdr[2].text = "Row"; hdr[3].text = "Description"
                for _, r in issues.iterrows():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(r["Field"])
                    row_cells[1].text = str(r["Issue Type"])
                    row_cells[2].text = str(r["Row"])
                    row_cells[3].text = str(r["Description"])
            word_doc.save(word_buffer)
            word_buffer.seek(0)

            st.download_button(
                label="📥 Download Word Report",
                data=word_buffer,
                file_name=f"Validation_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            # --- Excel Report ---
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                summary_df.to_excel(writer, sheet_name="Validation Issues", index=False)
            excel_buffer.seek(0)

            st.download_button(
                label="📊 Download Excel Report",
                data=excel_buffer,
                file_name=f"Validation_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        else:
            st.success("✅ No validation issues found!")
