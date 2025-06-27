import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import io

st.title("Data Quality Validation App")

# --- Step 1: Upload Files ---
st.header("Step 1: Upload Key and Data Files")
key_file = st.file_uploader("Upload the KEY Excel file", type=["xlsx"], key="key")
data_file = st.file_uploader("Upload the DATA Excel file", type=["xlsx"], key="data")

key_description, data_sheets, answer_map = None, None, {}
row_only_fields = []
row_or_column_fields = ['VisitType', 'ReNoSchool', 'RChildType', 'RHhType', 'EducStatus']

if key_file and data_file:
    key_description = pd.read_excel(key_file, sheet_name=None)
    data_sheets = pd.read_excel(data_file, sheet_name=None)

    answer_list_df = key_description.get('answer_list', pd.DataFrame())
    if not answer_list_df.empty:
        answer_list_df.dropna(how='all', inplace=True)
        answer_list_df.dropna(axis=1, how='all', inplace=True)
        grouped = answer_list_df.groupby('Field Name Eng for partner')
        for field, group in grouped:
            if pd.isna(field):
                continue
            all_values = []
            for _, row in group.iterrows():
                if field in row_only_fields:
                    values = row.drop(labels=['Language', 'Field Name Eng for partner']).dropna().tolist()
                elif field in row_or_column_fields:
                    row_vals = row.drop(labels=['Language', 'Field Name Eng for partner']).dropna().tolist()
                    col_vals = group.columns[2:][~row[2:].isna()].tolist()
                    values = row_vals + col_vals
                else:
                    values = group.columns[2:][~row[2:].isna()].tolist()
                all_values.extend([str(v).strip() for v in values if str(v).strip()])
            unique_values = list(dict.fromkeys(all_values))
            if unique_values:
                answer_map[field] = unique_values

    # Step 2: Country and Language Selection
    st.header("Step 2: Select Country and Language")
    country_choice = st.selectbox("Select Country", options=['GHA', 'CIV'], index=0)
    language_choice = st.selectbox("Select Language", options=['EN', 'FR'], index=0)

    # Step 3: Run Validation
    if st.button("Run Validation"):
        description_df = key_description['description'] if 'description' in key_description else key_description
        required_meta = description_df[description_df['For Mars KPI reporting'].isin(['Y', country_choice])]
        expected_fields = required_meta['Field Name Eng for partner'].dropna().unique().tolist()
        field_types = dict(zip(description_df['Field Name Eng for partner'], description_df['Type of variable in the table']))

        data_issues = []

        def log_issue(sheet, field, row, issue_type, description):
            data_issues.append({
                'Sheet': sheet, 'Field': field, 'Row': row,
                'Issue Type': issue_type, 'Description': description
            })

        def validate_sheet(df, sheet_name):
            validation_msgs = {
                'empty_sheet': {'EN': 'Sheet is present but contains no data', 'FR': 'La feuille est présente mais ne contient aucune donnée'},
                'completely_empty': {'EN': 'The variable "{}" is completely empty.', 'FR': 'La variable "{}" est complètement vide.'},
                'missing_required': {'EN': 'Field is required but missing', 'FR': 'Champ requis mais manquant'},
                'expected_date': {'EN': 'Expected format is DD-MM-YYYY', 'FR': 'Format attendu : JJ-MM-AAAA'},
                'expected_numeric': {'EN': 'Expected a numeric value', 'FR': 'Une valeur numérique était attendue'},
                'invalid_value': {'EN': "'{}' not in allowed values", 'FR': "'{}' ne fait pas partie des valeurs autorisées"},
                'alpha_numeric': {'EN': 'Value must be alphanumeric, not digits only', 'FR': 'La valeur doit être alphanumérique, pas uniquement des chiffres'},
                'duplicate_combo': {'EN': 'Duplicate entries found based on combination of: {}', 'FR': 'Doublons détectés selon la combinaison : {}'},
                'cond_re_no': {'EN': 'ChldAvble is 0, ReNoAvble must have a value', 'FR': 'ChldAvble est 0, ReNoAvble doit contenir une valeur'},
                'cond_hh_re_no': {'EN': 'HH_Avble is 0, HH_ReNoAvble must have a value', 'FR': 'HH_Avble est 0, HH_ReNoAvble doit contenir une valeur'}
            }

            if df.empty:
                log_issue(sheet_name, None, None, 'Empty Sheet', validation_msgs['empty_sheet'][language_choice])
                return

            actual_fields = df.columns.tolist()
            for field in expected_fields:
                if field not in actual_fields:
                    continue
                col_data = df[field].replace(r'^\s*$', np.nan, regex=True)

                # Conditional rule: ReNoAvble
                if sheet_name == 'D' and field == 'ReNoAvble' and 'ChldAvble' in df.columns:
                    for idx, val in col_data.items():
                        if str(df.loc[idx, 'ChldAvble']).strip() == '0' and pd.isna(val):
                            log_issue(sheet_name, field, idx, 'Conditional Rule', validation_msgs['cond_re_no'][language_choice])
                elif sheet_name == 'B-C' and field == 'HH_ReNoAvble' and 'HH_Avble' in df.columns:
                    for idx, val in col_data.items():
                        if str(df.loc[idx, 'HH_Avble']).strip() == '0' and pd.isna(val):
                            log_issue(sheet_name, field, idx, 'Conditional Rule', validation_msgs['cond_hh_re_no'][language_choice])

                if col_data.isnull().all():
                    log_issue(sheet_name, field, None, 'Missing Value', validation_msgs['completely_empty'][language_choice].format(field))
                    continue

                for idx, val in col_data.items():
                    if pd.isna(val):
                        log_issue(sheet_name, field, idx, 'Missing Value', validation_msgs['missing_required'][language_choice])
                        continue
                    if field in field_types:
                        expected_type = field_types[field]
                        if expected_type == 'datetime64[ns]':
                            try:
                                pd.to_datetime(val, format='%d-%m-%Y')
                            except:
                                log_issue(sheet_name, field, idx, 'Date Format Error', validation_msgs['expected_date'][language_choice])
                        elif expected_type == 'float64':
                            if not isinstance(val, (int, float, np.float64, np.int64)):
                                log_issue(sheet_name, field, idx, 'Type Error', validation_msgs['expected_numeric'][language_choice])
                    if field in answer_map:
                        val_norm = str(val).strip().lower()
                        allowed_vals = [str(v).strip().lower() for v in answer_map[field]]
                        if val_norm not in allowed_vals:
                            log_issue(sheet_name, field, idx, 'Invalid Value', validation_msgs['invalid_value'][language_choice].format(val))

            composite_fields = [f for f in ['ChldID', 'FarmerID', 'VisitType', 'EndDateActivity'] if f in df.columns]
            if len(composite_fields) >= 2:
                composite_key = df[composite_fields].astype(str).agg('|'.join, axis=1)
                dupes = composite_key[composite_key.duplicated()]
                if not dupes.empty:
                    combo = ', '.join(composite_fields)
                    log_issue(sheet_name, composite_fields[0], None, 'Duplicate', validation_msgs['duplicate_combo'][language_choice].format(combo))

        for sheet_name in ['P', 'B-C', 'D', 'E_Com', 'E_Hho', 'E_Chd']:
            if sheet_name in data_sheets:
                validate_sheet(data_sheets[sheet_name], sheet_name)

        st.subheader("Step 4: Validation Results")
        if data_issues:
            df_issues = pd.DataFrame(data_issues)
            st.dataframe(df_issues)

            grouped_summary = df_issues.groupby(['Sheet', 'Field', 'Issue Type']).agg(
                Issue_Count=('Description', 'count'),
                Example_Description=('Description', 'first')
            ).reset_index()

            if st.button("Generate Word Report"):
                word_doc = Document()
                word_doc.add_heading("MARS Data Quality Report", 0)

                for sheet in df_issues['Sheet'].unique():
                    word_doc.add_heading(f"Sheet: {sheet}", level=1)
                    issues = grouped_summary[grouped_summary['Sheet'] == sheet]
                    for _, row in issues.iterrows():
                        msg = f"[{row['Issue Type']}] {row['Field']}: {row['Example_Description']} (Total: {row['Issue_Count']})"
                        word_doc.add_paragraph(msg, style='ListBullet')

                filename = f"Validation_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                buffer = io.BytesIO()
                word_doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label="📥 Download Word Report",
                    data=buffer,
                    file_name=filename,
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
        else:
            st.success("No validation issues found!")
