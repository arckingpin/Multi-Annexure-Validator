import streamlit as st
import pandas as pd
import re
from io import BytesIO

# -------------------------------
# Helper: Convert DataFrame to Excel bytes
# -------------------------------
def to_excel_bytes(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name="Sheet1")
    writer.close()  # Use close() instead of save()
    processed_data = output.getvalue()
    return processed_data

# -------------------------------
# PAGE CONFIGURATION & CSS
# -------------------------------
st.set_page_config(
    page_title="Excel Validation App (Interactive Fixes)",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(
    """
    <style>
    /* Overall background */
    .reportview-container {
        background: #f0f2f6;
    }
    /* Header style */
    .css-18e3th9 {
        font-size: 2.5em;
        font-weight: 600;
        color: #333;
    }
    /* Sidebar styling */
    .sidebar .sidebar-content {
        background-color: #ffffff;
        border-right: 1px solid #e6e6e6;
    }
    /* Modern button style */
    .stButton>button {
        background-color: #0073e6;
        color: #fff;
        border: none;
        border-radius: 4px;
        padding: 8px 16px;
        font-size: 1em;
        transition: background-color 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #005bb5;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------
# SIDEBAR INSTRUCTIONS
# -------------------------------
st.sidebar.header("Instructions")
st.sidebar.markdown(
    """
    **Step 1: Upload Files**
    - Upload the **Validation Master** Excel file.  
      This file must contain two sheets:
        1. A sheet with the validation conditions table.  
           **IMPORTANT:** This sheet must have exactly 6 columns in the following order:
              1. Field code  
              2. Field Name  
              3. Data Type  
              4. Validation  
              5. Whether Optional/Mandatory  
              6. Description
        2. A sheet with valid state names (the first column holds the valid state names).
    - Upload the **Input Excel** file to be validated.
    
    **Step 2: Sheet Selection**
    - Choose which sheet holds the validation rules and which holds the state master.
    
    **Step 3: Validation**
    - The app first validates the input Excel against the master rules.
    - Additionally, independent of the master, for each input column:
         - If the column header contains the word “time” (case‑insensitive) then the data is expected to be datetime with format **dd-mm-yyyy hh:mm**.
         - If the column header contains “date” (but not “time”) then the data is expected to be a date with format **dd-mm-yyyy**.
    - For any field that fails conversion, a fixable error is shown with a drop‑down letting you select a new data type.
      If you select “date”, a text input appears with a default date format (automatically set based on the header).
      When you click **Submit Fix**, the column is updated accordingly.
    
    **Step 4: Download**
    - Once fixes are applied, click **Download Modified Excel** to download the updated file.
    """
)

# -------------------------------
# FILE UPLOADS
# -------------------------------
st.title("Excel Validation App (Interactive Fixes)")
st.markdown("#### Upload the files for validation")

validation_master_file = st.file_uploader("Upload Validation Master File", type=["xlsx"], key="validation_master")
input_excel_file = st.file_uploader("Upload Input Excel File", type=["xlsx"], key="input_excel")

if validation_master_file and input_excel_file:
    try:
        # Load the Validation Master file and list available sheets
        xls = pd.ExcelFile(validation_master_file)
        sheet_names = xls.sheet_names

        st.markdown("### Choose Sheets from Validation Master")
        selected_validation_sheet = st.selectbox("Select the sheet for Validation Rules", sheet_names, key="val_sheet")
        selected_state_sheet = st.selectbox("Select the sheet for State Master", sheet_names, key="state_sheet")

        # Read the selected validation sheet
        df_validations = pd.read_excel(xls, sheet_name=selected_validation_sheet)
        # Ensure exactly 6 columns (ignoring header names)
        if df_validations.shape[1] != 6:
            st.error("The selected validation sheet must contain exactly 6 columns (in the correct order).")
            st.stop()
        # Rename columns by fixed order regardless of the file's header names
        df_validations.columns = [
            "Field code", 
            "Field Name", 
            "Data Type", 
            "Validation", 
            "Whether Optional/Mandatory", 
            "Description"
        ]
        
        # Read the selected state master sheet
        df_states = pd.read_excel(xls, sheet_name=selected_state_sheet)
        
        # ---------------------------------------------------------
        # BUILD VALIDATION SPECIFICATION DICTIONARY FROM MASTER
        # ---------------------------------------------------------
        # Use "Field Name" as key (which should match the Input Excel's headers)
        validation_spec = {}
        for _, row in df_validations.iterrows():
            field_name = row["Field Name"]
            data_type = row["Data Type"]
            mandatory_flag = row["Whether Optional/Mandatory"]
            mandatory = True if str(mandatory_flag).strip().upper() == "M" else False
            validation_rule = row["Validation"]
            description = row["Description"]
            validation_spec[field_name] = {
                "data_type": data_type,
                "mandatory": mandatory,
                "validation": validation_rule,
                "description": description
            }
        expected_fields = list(validation_spec.keys())
        
        # ---------------------------------------------------------
        # BUILD VALID STATE NAMES LIST
        # ---------------------------------------------------------
        valid_states = df_states.iloc[:, 0].astype(str).str.strip().tolist()
        
        st.subheader("Validation Master Preview")
        st.write("**Validation Rules (first few rows):**")
        st.dataframe(df_validations.head())
        st.write("**State Master (first few rows):**")
        st.dataframe(df_states.head())
        
        # ---------------------------------------------------------
        # READ AND PREVIEW THE INPUT EXCEL FILE
        # ---------------------------------------------------------
        df_input = pd.read_excel(input_excel_file)
        if "df_input" not in st.session_state:
            st.session_state.df_input = df_input.copy()
        df_input = st.session_state.df_input  # work with the session_state copy
        
        with st.expander("Preview Input Excel Data"):
            st.dataframe(df_input.head())
        
        # -------------------------------
        # VALIDATION PROCESS BASED ON MASTER
        # -------------------------------
        fixable_errors = {}   # for fields with fixable errors (e.g. date conversion errors)
        non_fixable_errors = []   # non-fixable errors
        
        # 1. Check for missing required columns (by Field Name from master)
        missing_fields = [field for field in expected_fields if field not in df_input.columns]
        if missing_fields:
            non_fixable_errors.append("Missing columns: " + ", ".join(missing_fields))
        
        # 2. Validate each field per the master specification
        for field, spec in validation_spec.items():
            if field in df_input.columns:
                # Mandatory check
                if spec["mandatory"] and df_input[field].isnull().any():
                    non_fixable_errors.append(f"Field '{field}' is mandatory but contains missing values.")
                # Data type validation per master
                expected_type = str(spec["data_type"]).lower()
                if expected_type == "number":
                    try:
                        pd.to_numeric(df_input[field])
                    except Exception:
                        non_fixable_errors.append(f"Field '{field}' should be numeric.")
                elif expected_type == "date":
                    try:
                        pd.to_datetime(df_input[field], errors='raise')
                    except Exception:
                        fixable_errors[field] = f"Field '{field}' should be a date (per master specification)."
                # Custom regex validation (if specified)
                if spec["validation"] and isinstance(spec["validation"], str) and spec["validation"].lower().startswith("regex:"):
                    pattern = spec["validation"][6:].strip()
                    invalid_entries = df_input[~df_input[field].astype(str).str.fullmatch(pattern, na=False)]
                    if not invalid_entries.empty:
                        non_fixable_errors.append(f"Field '{field}' does not match the expected pattern: {pattern}.")
        
        # -------------------------------
        # INDEPENDENT VALIDATION BASED ON COLUMN HEADER
        # -------------------------------
        # For every column in the input Excel, if the header contains "time" (case-insensitive),
        # enforce that the column values are convertible using dd-mm-yyyy hh:mm.
        # If the header contains "date" (but not "time"), enforce dd-mm-yyyy.
        for col in df_input.columns:
            lower_col = col.lower()
            desired_format = None
            if "time" in lower_col:
                desired_format = "%d-%m-%Y %H:%M"
            elif "date" in lower_col and "time" not in lower_col:
                desired_format = "%d-%m-%Y"
            if desired_format:
                try:
                    pd.to_datetime(df_input[col], format=desired_format, errors='raise')
                except Exception:
                    fixable_errors[col] = f"Column '{col}' should be a date/time with format {desired_format} (independent rule)."
        
        # -------------------------------
        # DISPLAY VALIDATION RESULTS
        # -------------------------------
        st.markdown("#### Validation Results")
        if non_fixable_errors or fixable_errors:
            if non_fixable_errors:
                st.error("The following validation errors were found:")
                for err in non_fixable_errors:
                    st.write("• " + err)
            if fixable_errors:
                st.warning("The following fields have errors that can be fixed:")
                # For each fixable error (date/datetime conversion errors)
                for field, err_msg in fixable_errors.items():
                    with st.expander(f"Fix error for field '{field}'", expanded=True):
                        st.write(err_msg)
                        new_type = st.selectbox(f"Select new data type for '{field}'", 
                                                  options=["string", "number", "date"],
                                                  index=2, key=f"{field}_new_type")
                        date_format = ""
                        if new_type == "date":
                            # Set default based on header: if header contains "time" use datetime format, else use date format.
                            default_format = "%d-%m-%Y %H:%M" if "time" in field.lower() else "%d-%m-%Y"
                            date_format = st.text_input(f"Enter date format for '{field}'", value=default_format, key=f"{field}_date_format")
                        if st.button(f"Submit Fix for {field}", key=f"submit_fix_{field}"):
                            df = st.session_state.df_input.copy()
                            try:
                                if new_type == "date":
                                    df[field] = pd.to_datetime(df[field], format=date_format, errors='coerce')
                                elif new_type == "number":
                                    df[field] = pd.to_numeric(df[field], errors='coerce')
                                elif new_type == "string":
                                    df[field] = df[field].astype(str)
                                st.success(f"Field '{field}' fixed successfully!")
                                st.session_state.df_input = df
                                if hasattr(st, "experimental_rerun"):
                                    st.experimental_rerun()
                                else:
                                    st.info("Please refresh the page to see updated data.")
                            except Exception as e:
                                st.error(f"Error fixing field '{field}': {e}")
        else:
            st.success("Congratulations! The input Excel file is valid as per the provided validation master and header rules.")
        
        # -------------------------------
        # DOWNLOAD MODIFIED EXCEL BUTTON
        # -------------------------------
        st.markdown("#### Download Modified Excel")
        modified_df = st.session_state.df_input.copy()
        excel_bytes = to_excel_bytes(modified_df)
        st.download_button(
            label="Download Modified Excel",
            data=excel_bytes,
            file_name="modified_input.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error("Error processing files: " + str(e))
