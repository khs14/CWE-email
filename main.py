import streamlit as st
import pandas as pd
import dns.resolver
import smtplib
import re
import socket
import concurrent.futures


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
        answers = dns.resolver.resolve(domain, 'MX')
        return True
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.exception.Timeout, socket.gaierror, UnicodeError):
        return False


def smtp_check(email):
    domain = email.split('@')[1]
    try:
        # Increase timeout for DNS resolution
        mx_records = dns.resolver.resolve(domain, 'MX', lifetime=20)
        mx_record = mx_records[0].exchange.to_text()

        # Connect to the SMTP server
        # Increase timeout for SMTP connection
        server = smtplib.SMTP(timeout=20)
        server.set_debuglevel(0)

        server.connect(mx_record)
        server.helo(server.local_hostname)
        server.mail('test@example.com')
        code, message = server.rcpt(email)
        server.quit()

        # 250 is the code for successful recipient
        if code == 250:
            return True
        else:
            return False
    except (smtplib.SMTPConnectError, smtplib.SMTPServerDisconnected, smtplib.SMTPResponseException, dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.exception.Timeout, socket.gaierror, UnicodeError):
        return False


def validate_email(email):
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

    duplicates = df_cleaned[df_cleaned.duplicated(subset=columns, keep=False)]
    # Sort by the selected columns to group duplicates together
    duplicates_sorted = duplicates.sort_values(by=columns)
    return duplicates_sorted


def find_missing_values(df, columns):
    # Find rows with None or NaN values in the specified columns
    missing_values = df[df[columns].isnull().any(axis=1)]
    return missing_values


st.title("CWE Data Tools")

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Uploaded DataFrame:")
    st.write(df)

    # Find the email column
    email_column = st.selectbox(
        "Select the column containing email addresses", df.columns)

    if email_column:
        # Add a button to start the validation process
        if st.button("Validate Emails"):
            email_list = df[email_column].tolist()

            invalid_emails = []
            progress_text = st.empty()
            progress_bar = st.progress(0)

            # Use ThreadPoolExecutor for concurrent email validation
            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                futures = {executor.submit(
                    validate_email, email): email for email in email_list}
                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    result = future.result()
                    if result is not None:
                        invalid_emails.append(result)
                    progress_text.text(
                        f"Processing: {i+1}/{len(email_list)} emails")
                    progress_bar.progress((i+1) / len(email_list))

            # Display the invalid emails
            if invalid_emails:
                st.write("Invalid Emails:")
                for email, reason in invalid_emails:
                    st.write(f"The email '{email}' is invalid: {reason}")
            else:
                st.write("All emails are valid.")

        # Multi-select box for selecting columns to check for duplicates
        columns_to_check = st.multiselect(
            "Select columns to check for duplicate values", df.columns)

        if columns_to_check:
            # Add a button to check for duplicates
            if st.button("Check Duplicates"):
                duplicates = find_duplicates(df, columns_to_check)
                if not duplicates.empty:
                    st.write("Duplicate Rows:")
                    st.write(duplicates)
                else:
                    st.write("No duplicate rows found.")

        # Multi-select box for selecting columns to check for missing values
        columns_to_check_missing = st.multiselect(
            "Select columns to check for missing values", df.columns)

        if columns_to_check_missing:
            # Add a button to check for missing values
            if st.button("Check Missing Values"):
                missing_values = find_missing_values(
                    df, columns_to_check_missing)
                if not missing_values.empty:
                    st.write("Rows with Missing Values:")
                    st.write(missing_values)
                else:
                    st.write("No missing values found in the selected columns.")
