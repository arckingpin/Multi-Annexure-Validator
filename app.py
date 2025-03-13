import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import io

def validate_and_fix_data(input_df, validation_rules, state_master):
    errors = []
    modified_df = input_df.copy()
    
    for index, row in validation_rules.iterrows():
        field_name = row[1]  # 2nd column is field name
        data_type = row[2]   # 3rd column is data type
        mandatory = row[4]   # 5th column is Mandatory (M) or Optional (O)
        
        if field_name in input_df.columns:
            column_data = input_df[field_name]
            
            if 'date' in field_name.lower() and 'time' not in field_name.lower():
                expected_format = '%d-%m-%Y'
                try:
                    modified_df[field_name] = pd.to_datetime(column_data, errors='coerce').dt.strftime(expected_format)
                except:
                    errors.append(f"Field '{field_name}' should be a date.")
            elif 'time' in field_name.lower():
                expected_format = '%d-%m-%Y %H:%M'
                try:
                    modified_df[field_name] = pd.to_datetime(column_data, errors='coerce').dt.strftime(expected_format)
                except:
                    errors.append(f"Field '{field_name}' should be a datetime.")
    
    return modified_df, errors

def main():
    st.title("Excel Validation & Fixing Tool")
    
    validation_file = st.file_uploader("Upload Validation Master", type=['xlsx'])
    input_file = st.file_uploader("Upload Input Excel", type=['xlsx'])
    
    if validation_file and input_file:
        validation_xl = pd.ExcelFile(validation_file)
        input_xl = pd.ExcelFile(input_file)
        
        sheet_options = validation_xl.sheet_names
        validation_sheet = st.selectbox("Select Validation Sheet", sheet_options)
        state_master_sheet = st.selectbox("Select State Master Sheet", sheet_options)
        
        validation_rules = validation_xl.parse(validation_sheet).iloc[:, :6]
        state_master = validation_xl.parse(state_master_sheet).iloc[:, 0]
        input_df = input_xl.parse(input_xl.sheet_names[0])
        
        modified_df, errors = validate_and_fix_data(input_df, validation_rules, state_master)
        
        for error in errors:
            col_name = error.split("'")[1]
            st.error(error)
            if st.button(f"Fix '{col_name}'"):
                original_data = input_df[col_name].astype(str).head(5)
                modified_data = modified_df[col_name].astype(str).head(5)
                
                st.write("### Before Fixing:")
                st.dataframe(original_data)
                
                st.write("### After Fixing:")
                st.dataframe(modified_data)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            modified_df.to_excel(writer, index=False, sheet_name='Validated_Data')
        
        st.download_button("Download Modified Excel", data=output.getvalue(), file_name="validated_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
if __name__ == "__main__":
    main()
