import streamlit as st
import pandas as pd
import dns.resolver
import smtplib
import re
import socket
from io import BytesIO
import time

# Email validation functions


def is_valid_email_format(email):
    if not isinstance(email, str):
        return False
    email = email.strip()  # Remove leading/trailing whitespace
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email) is not None


def has_valid_mx_record(domain):
    try:
        if not domain or len(domain) > 255:
            return False
        dns.resolver.resolve(domain, 'MX')
        return True
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.exception.Timeout, socket.gaierror, UnicodeError):
        return False


def smtp_check(email):
    domain = email.split('@')[1]
    try:
        mx_records = dns.resolver.resolve(domain, 'MX', lifetime=20)
        mx_record = mx_records[0].exchange.to_text()
        server = smtplib.SMTP(timeout=20)
        server.set_debuglevel(0)
        server.connect(mx_record)
        server.helo(server.local_hostname)
        server.mail('test@example.com')
        code, _ = server.rcpt(email)
        server.quit()
        return code == 250
    except (smtplib.SMTPConnectError, smtplib.SMTPServerDisconnected, smtplib.SMTPResponseException, dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.exception.Timeout, socket.gaierror, UnicodeError, smtplib.SMTPException) as e:
        st.write(f"SMTP check failed for {email}: {e}")
        return False


def validate_email(email):
    if email is None:
        return (email, "Email is None")
    if is_valid_email_format(email):
        domain = email.split('@')[1]
        if not has_valid_mx_record(domain):
            return (email, "Invalid MX record")
        elif not smtp_check(email):
            return (email, "SMTP check failed")
    else:
        return (email, "Invalid format")
    return None


def find_duplicates(df_cleaned, columns):
    df_temp = df_cleaned.copy()
    df_temp[columns] = df_temp[columns].apply(
        lambda x: x.str.strip() if x.dtype == "object" else x)
    duplicates = df_temp[df_temp.duplicated(subset=columns, keep=False)]
    return duplicates.sort_values(by=columns)


def find_missing_values(df, columns):
    return df[df[columns].isnull().any(axis=1)]


def process_email_batch(email_list):
    invalid_emails = []
    for email in email_list:
        try:
            result = validate_email(email)
            if result is not None:
                invalid_emails.append(result)
        except Exception as e:
            invalid_emails.append((email, f"Error: {e}"))
        time.sleep(2)  # Sleep for 2 seconds between each email check
    return invalid_emails


def convert_df_to_excel(dfs, sheetnames):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheetname in zip(dfs, sheetnames):
            df.to_excel(writer, index=False, sheet_name=sheetname)
    return output.getvalue()


# Streamlit UI
st.title("Data Cleaning Tool")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Uploaded DataFrame:")
    st.write(df)

    # Email Validation Section
    st.header("Email Validator")
    email_column = st.selectbox(
        "Select the column containing email addresses", df.columns, key="email")

    if email_column:
        if st.button("Validate Emails"):
            email_list = df[email_column].tolist()

            invalid_emails = []
            progress_text = st.empty()
            progress_bar = st.progress(0)

            result_container = st.empty()
            total_emails = len(email_list)

            for i, email in enumerate(email_list):
                batch_invalid_emails = process_email_batch([email])
                invalid_emails.extend(batch_invalid_emails)
                progress_text.text(f"Processing: {i + 1}/{total_emails}")
                progress_bar.progress((i + 1) / total_emails)

                if batch_invalid_emails:
                    result_container.write("Invalid Emails:")
                    for email, reason in batch_invalid_emails:
                        result_container.write(
                            f"The email '{email}' is invalid: {reason}")

            if invalid_emails:
                invalid_emails_df = pd.DataFrame(
                    invalid_emails, columns=['Email', 'Reason'])
                excel_data = convert_df_to_excel(
                    [invalid_emails_df], ["Invalid Emails"])
                st.download_button(
                    label="Download Excel of Invalid Emails",
                    data=excel_data,
                    file_name='invalid_emails.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

    # Duplicate Check Section
    st.header("Duplicate Checker")
    columns_to_check = st.multiselect(
        "Select columns to check for duplicate values", df.columns, key="duplicates")

    if columns_to_check:
        if st.button("Check Duplicates"):
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

    # Missing Value Check Section
    st.header("Missing Value Finder")
    columns_to_check_missing = st.multiselect(
        "Select columns to check for missing values", df.columns, key="missing")

    if columns_to_check_missing:
        if st.button("Check Missing Values"):
            missing_values = find_missing_values(df, columns_to_check_missing)
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
                st.write("No missing values found in the selected columns.")
