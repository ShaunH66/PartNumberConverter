import streamlit as st
import pandas as pd
import io

# --- Helper Functions ---

def to_excel(df):
    """Converts a pandas DataFrame to an in-memory Excel file (bytes)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='ConvertedParts')
    processed_data = output.getvalue()
    return processed_data

def load_data_file(uploaded_file, header_row=1):
    """
    Intelligently loads data files, allowing the user to specify the header row.
    Handles various Excel formats and delimited text files.
    """
    skiprows = header_row - 1
    file_name_lower = uploaded_file.name.lower()

    try:
        if file_name_lower.endswith(('.csv', '.tsv', '.txt')):
            return pd.read_csv(uploaded_file, on_bad_lines='skip', header=skiprows, sep=None, engine='python')

        elif file_name_lower.endswith('.xlsb'):
            return pd.read_excel(uploaded_file, engine='pyxlsb', header=skiprows)

        elif file_name_lower.endswith(('.xlsx', '.xlsm', '.xls')):
            try:
                return pd.read_excel(uploaded_file, engine='openpyxl', header=skiprows)
            except Exception as e:
                if "zip file" in str(e).lower() or file_name_lower.endswith('.xls'):
                    uploaded_file.seek(0)
                    return pd.read_excel(uploaded_file, engine='xlrd', header=skiprows)
                else:
                    raise e
    except Exception as e:
        st.error(f"Failed to read file '{uploaded_file.name}'. Error: {e}")
        st.info("Please ensure the 'Header is on which row?' value is correct. The selected row must contain the column names.")
        return None

# --- Streamlit App UI ---

st.set_page_config(layout="wide", page_title="Part Number Converter")
st.title("⚙️ MSE Part Number Conversion Tool")

with st.expander("ℹ️ Need Help? Click here for instructions and tips.", expanded=False):
    st.markdown("""
        This tool automates the process of replacing old part numbers with new ones by looking them up in a master reference file.

        ### How to Use
        
        **1. Upload Your Files:**
        -   **Master File:** This is your "answer key" or "lookup table". It must contain a column for the old part numbers (`E-Number`) and a column for the new part numbers (`200 Number`).
        -   **Data File:** This is the file you want to process. It contains a list of old part numbers that you need to convert.

        **2. Set the Header Row:**
        -   Look at your Excel or CSV file and find the row number where the column titles (like 'E Number', 'Description', etc.) are located.
        -   Enter this number into the **"Header is on which row?"** box for each file.
        -   *Example:* In your `Common E Numbers.xlsx` file, the headers are on **row 8**. You must enter `8` for the Master File. For most standard files, this will be `1`.

        **3. Select Key Columns:**
        -   **In the Master File:**
            -   **Select the 'E Number' column (the Key):** This is the column of old part numbers that will be used for matching.
            -   **Select the 'New Part Number' column (the Value):** This is the column of new part numbers that you want to get.
        -   **In your Data File:**
            -   **Select the column with E-Numbers to convert:** Point to the column in your data file that contains the old part numbers.

        **4. Convert and Download:**
        -   Click the **"Convert Part Numbers"** button. The app will add a new column to your data file with the converted numbers.
        -   You can then download the result as a new Excel file.

        ### Troubleshooting Tips

        *   **My columns are named `Unnamed: 0`, `Unnamed: 1`, etc.**
            -   This means the **"Header is on which row?"** number is incorrect. Double-check your file in Excel and set the correct row number.
        *   **I see `--- NOT FOUND ---` in the results.**
            -   This is expected. It means an E-Number from your Data File did not exist in the Master File's lookup column.
        *   **The app shows an error like "Failed to read file".**
            -   First, check that the Header Row number is correct. If it is, the file might be corrupted or in a rare, unsupported format. Try re-saving it in Excel as a standard `.xlsx` file.
    """)

# --- 1. File Uploaders ---
st.header("1. Upload Your Files")

SUPPORTED_TYPES = ["csv", "tsv", "txt", "xlsx", "xls", "xlsm", "xlsb"]
col1, col2 = st.columns(2)

with col1:
    master_file = st.file_uploader("Upload Master File", type=SUPPORTED_TYPES)
    master_header_row = st.number_input("Header is on which row in Master File?", min_value=1, step=1, value=8)

with col2:
    user_file = st.file_uploader("Upload Your Data File", type=SUPPORTED_TYPES)
    user_header_row = st.number_input("Header is on which row in Data File?", min_value=1, step=1, value=1)

# --- 2. Load data and show UI ---
if master_file and user_file:
    with st.spinner("Loading files..."):
        master_df = load_data_file(master_file, master_header_row)
        user_df = load_data_file(user_file, user_header_row)

    if master_df is not None and user_df is not None:
        st.success("Files loaded successfully! Please select the columns below.")

        with st.expander("Show Master File Preview", expanded=True):
            st.dataframe(master_df.head())
        with st.expander("Show Your Data File Preview", expanded=True):
            st.dataframe(user_df.head())

        st.header("2. Select Key Columns")
        master_cols = master_df.columns.tolist()
        user_cols = user_df.columns.tolist()
        col1_select, col2_select = st.columns(2)

        with col1_select:
            st.subheader("In your Master File:")
            master_e_col_index = master_cols.index('E Number') if 'E Number' in master_cols else 0
            new_part_col_index = master_cols.index('200 Number') if '200 Number' in master_cols else 1
            master_e_col = st.selectbox("Select the 'E Number' column (the Key):", master_cols, index=master_e_col_index)
            new_part_col = st.selectbox("Select the 'New Part Number' column (the Value):", master_cols, index=new_part_col_index)

        with col2_select:
            st.subheader("In your Data File:")
            user_e_col_index = user_cols.index('PART/ E #') if 'PART/ E #' in user_cols else 0
            user_e_col = st.selectbox("Select the column with E-Numbers to convert:", user_cols, index=user_e_col_index)

        st.header("3. Process and Download")

        if st.button("🚀 Convert Part Numbers", type="primary"):
            with st.spinner("Processing..."):
                user_df_copy = user_df.copy()
                master_df_copy = master_df.copy()

                mapping_df = master_df_copy[[master_e_col, new_part_col]].copy()
                mapping_df[master_e_col] = mapping_df[master_e_col].astype(str).str.strip()
                user_df_copy[user_e_col] = user_df_copy[user_e_col].astype(str).str.strip()
                mapping_df.drop_duplicates(subset=[master_e_col], inplace=True)

                result_df = pd.merge(user_df_copy, mapping_df, left_on=user_e_col, right_on=master_e_col, how='left')
                
                new_col_name = "Converted Part Number"
                result_df.rename(columns={new_part_col: new_col_name}, inplace=True)
                if user_e_col != master_e_col:
                    result_df.drop(columns=[master_e_col], inplace=True)
                result_df[new_col_name].fillna('--- NOT FOUND ---', inplace=True)

                st.success("✅ Conversion Complete!")
                st.dataframe(result_df)
                
                excel_data = to_excel(result_df)
                st.download_button(
                    label="📥 Download Converted Excel File",
                    data=excel_data,
                    file_name="converted_part_numbers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("---")