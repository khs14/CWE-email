import streamlit as st
import pandas as pd
import dns.resolver
import smtplib
import re
import socket
from io import BytesIO
import time


def is_valid_email_format(email):
    if not isinstance(email, str):
        return False
    
    email = email.strip().lower()
    
    # Check length
    if len(email) > 254:
        return False
    
    # Split local part and domain
    try:
        local_part, domain = email.split('@')
    except ValueError:
        return False
    
    # Check local part
    if len(local_part) > 64 or len(local_part) == 0:
        return False
    
    # Check domain
    if len(domain) > 253 or len(domain) == 0:
        return False
    
    # Comprehensive regex for email validation
    pattern = r'^[a-z0-9!#$%&\'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&\'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$'
    
    return re.match(pattern, email) is not None

def find_duplicates(df_cleaned, columns):
    df_temp = df_cleaned.copy()
    df_temp[columns] = df_temp[columns].apply(
        lambda x: x.str.strip() if x.dtype == "object" else x)
    duplicates = df_temp[df_temp.duplicated(subset=columns, keep=False)]
    return duplicates.sort_values(by=columns)


def find_missing_values(df, columns):
    missing_patterns = ['null', 'na', 'none', '--']

    def is_missing(value):
        if pd.isnull(value):
            return True
        elif isinstance(value, str) and value.lower() in missing_patterns:
            return True
        else:
            return False

    missing_values_mask = df[columns].applymap(is_missing)
    missing_values = df[missing_values_mask.any(axis=1)]
    return missing_values


def convert_df_to_excel(dfs, sheetnames):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheetname in zip(dfs, sheetnames):
            df.to_excel(writer, index=False, sheet_name=sheetname)
    return output.getvalue()


# Streamlit UI
st.title("CWE data cleaning tool")

# Sidebar for navigation
page = st.sidebar.selectbox("Choose a function", [
                            "Email Validator", "Duplicate Checker", "Missing Value Finder", "Compare Excel Files"])



# Email validation functions

def is_valid_email_format(email):
    if not isinstance(email, str):
        return False
    email = email.strip()  # Remove leading/trailing whitespace
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email) is not None

if st.button("Validate Emails"):
    email_list = df[email_column].tolist()

    invalid_emails = []
    progress_text = st.empty()
    progress_bar = st.progress(0)

    result_container = st.empty()
    total_emails = len(email_list)

    for i, email in enumerate(email_list):
        result = validate_email(email)
        if result is not None:
            invalid_emails.append(result)

        progress_text.text(f"Checking: {i + 1}/{total_emails}")
        progress_bar.progress((i + 1) / total_emails)

    if invalid_emails:
        result_container.write("Validation Results:")
        for email, status in invalid_emails:
            result_container.write(f"Email: {email}, Status: {status}")

        # Create DataFrame for download
        results_df = pd.DataFrame(invalid_emails, columns=['Email', 'Status'])
        excel_data = convert_df_to_excel([results_df], ["Validation Results"])
        st.download_button(
            label="Download Validation Results",
            data=excel_data,
            file_name='email_validation_results.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        result_container.write("All emails passed the validation.")

elif page == "Duplicate Checker":
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select a sheet", xls.sheet_names)

        if sheet_name:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            st.write(f"Data from {sheet_name}:")
            st.write(df)

            st.header("Duplicate Checker")
            columns_to_check = st.multiselect(
                "Select columns to check for duplicate values", df.columns, key="duplicates")

            if columns_to_check:
                if st.button("Check Duplicates"):
                    def find_duplicates(df, columns):
                        # Create a copy of the dataframe to avoid modifying the original
                        df_clean = df.copy()

                        # Define missing patterns
                        missing_patterns = [
                            'null', 'na', 'none', '--', '', 'nan']

                        # Clean the data in the selected columns
                        for col in columns:
                            df_clean[col] = df_clean[col].astype(
                                str).str.strip().str.lower()
                            df_clean[col] = df_clean[col].replace(
                                missing_patterns, pd.NA)

                        # Remove rows where any of the selected columns have NA values
                        df_clean = df_clean.dropna(subset=columns)

                        # Find and sort duplicates
                        duplicates = df_clean[df_clean.duplicated(
                            subset=columns, keep=False)]
                        duplicates = duplicates.sort_values(by=columns)

                        return duplicates

                    duplicates = find_duplicates(df, columns_to_check)
                    if not duplicates.empty:
                        st.write("Duplicate Rows:")
                        st.write(duplicates)
                        duplicates_df = duplicates.copy()
                        excel_data = convert_df_to_excel(
                            [duplicates_df], ["Duplicates"])
                        st.download_button(
                            label="Download Excel of Duplicate Rows",
                            data=excel_data,
                            file_name='duplicates.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    else:
                        st.write("No duplicate rows found.")


elif page == "Missing Value Finder":
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Select a sheet", xls.sheet_names)

        if sheet_name:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            st.write(f"Data from {sheet_name}:")
            st.write(df)

            st.header("Missing Value Finder")
            columns_to_check_missing = st.multiselect(
                "Select columns to check for missing values", df.columns, key="missing")

            if columns_to_check_missing:
                if st.button("Check Missing Values"):
                    missing_values = find_missing_values(
                        df, columns_to_check_missing)
                    if not missing_values.empty:
                        st.write("Rows with Missing Values:")
                        st.write(missing_values)
                        missing_values_df = missing_values.copy()
                        excel_data = convert_df_to_excel(
                            [missing_values_df], ["Missing Values"])
                        st.download_button(
                            label="Download Excel of Missing Values",
                            data=excel_data,
                            file_name='missing_values.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    else:
                        st.write(
                            "No missing values found in the selected columns.")

elif page == "Compare Excel Files":
    uploaded_file1 = st.file_uploader(
        "Choose the first Excel file", type="xlsx")
    uploaded_file2 = st.file_uploader(
        "Choose the second Excel file", type="xlsx")

    if uploaded_file1 is not None and uploaded_file2 is not None:
        xls1 = pd.ExcelFile(uploaded_file1)
        xls2 = pd.ExcelFile(uploaded_file2)

        sheet_names1 = xls1.sheet_names
        sheet_names2 = xls2.sheet_names

        col1, col2 = st.columns(2)

        with col1:
            sheet_name1 = st.selectbox(
                "Select a sheet from the first Excel file", sheet_names1)
            if sheet_name1:
                df1 = pd.read_excel(xls1, sheet_name=sheet_name1)
                st.write(f"Data from {sheet_name1} (First Excel):")
                st.write(df1)
                column1 = st.selectbox(
                    "Select column from first Excel", df1.columns)

        with col2:
            sheet_name2 = st.selectbox(
                "Select a sheet from the second Excel file", sheet_names2)
            if sheet_name2:
                df2 = pd.read_excel(xls2, sheet_name=sheet_name2)
                st.write(f"Data from {sheet_name2} (Second Excel):")
                st.write(df2)
                column2 = st.selectbox(
                    "Select column from second Excel", df2.columns)

        if sheet_name1 and sheet_name2 and column1 and column2:
            if st.button("Compare Sheets"):
                def clean_and_compare_dataframes(df1, df2, column1, column2):
                    missing_patterns = ['null', 'na', 'none', '--']

                    # Clean data: remove extra whitespace, convert to lowercase, and handle missing values
                    def clean_data(df, column):
                        df[column] = df[column].astype(
                            str).str.strip().str.lower()
                        df[column] = df[column].replace(
                            missing_patterns, pd.NA)
                        return df.dropna(subset=[column])

                    df1_clean = clean_data(df1, column1)
                    df2_clean = clean_data(df2, column2)

                    df1_temp = df1_clean.rename(columns={column1: 'merge_key'})
                    df2_temp = df2_clean.rename(columns={column2: 'merge_key'})

                    common_rows = pd.merge(df1_temp, df2_temp, on='merge_key')
                    diff_rows_df1 = df1_temp[~df1_temp['merge_key'].isin(
                        df2_temp['merge_key'])]
                    diff_rows_df2 = df2_temp[~df2_temp['merge_key'].isin(
                        df1_temp['merge_key'])]

                    return common_rows, diff_rows_df1, diff_rows_df2

                common_rows, diff_rows_df1, diff_rows_df2 = clean_and_compare_dataframes(
                    df1, df2, column1, column2)

                st.write("Common Rows:")
                st.write(common_rows)

                st.write(f"Rows in {sheet_name1} but not in {sheet_name2}:")
                st.write(diff_rows_df1)

                st.write(f"Rows in {sheet_name2} but not in {sheet_name1}:")
                st.write(diff_rows_df2)

                def truncate_sheet_name(name, max_length=31):
                    if len(name) <= max_length:
                        return name
                    return name[:max_length-3] + '...'

                sheet1_name = truncate_sheet_name(f"Only in {sheet_name1}")
                sheet2_name = truncate_sheet_name(f"Only in {sheet_name2}")

                # Prepare Excel files for download
                common_excel = convert_df_to_excel(
                    [common_rows], ["Common Rows"])
                diff1_excel = convert_df_to_excel(
                    [diff_rows_df1], [sheet1_name])
                diff2_excel = convert_df_to_excel(
                    [diff_rows_df2], [sheet2_name])
                all_results_excel = convert_df_to_excel(
                    [common_rows, diff_rows_df1, diff_rows_df2],
                    ["Common Rows", sheet1_name, sheet2_name]
                )

                # Create download buttons
                button_col1, button_col2 = st.columns(2)

                with button_col1:
                    st.download_button(
                        label="Download Common Rows",
                        data=common_excel,
                        file_name='common_rows.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                    st.download_button(
                        label=f"Download Rows Only in {sheet_name1}",
                        data=diff1_excel,
                        file_name=f'only_in_{sheet_name1}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                with button_col2:
                    st.download_button(
                        label=f"Download Rows Only in {sheet_name2}",
                        data=diff2_excel,
                        file_name=f'only_in_{sheet_name2}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                    st.download_button(
                        label="Download All Results",
                        data=all_results_excel,
                        file_name='all_comparison_results.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

    else:
        st.info("Please upload both Excel files to compare.")
