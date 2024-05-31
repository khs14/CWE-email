import streamlit as st
import pandas as pd
import dns.resolver
import smtplib
import re
import socket


def is_valid_email_format(email):
    if not isinstance(email, str):
        return False
    email = email.strip()  # Remove leading/trailing whitespace
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email) is not None


def has_valid_mx_record(domain):
    try:
        answers = dns.resolver.resolve(domain, 'MX')
        return True
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.exception.Timeout, socket.gaierror):
        return False


def smtp_check(email):
    domain = email.split('@')[1]
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        mx_record = mx_records[0].exchange.to_text()

        # Connect to the SMTP server
        server = smtplib.SMTP(timeout=10)
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
    except (smtplib.SMTPConnectError, smtplib.SMTPServerDisconnected, smtplib.SMTPResponseException, dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.exception.Timeout, socket.gaierror):
        return False


st.title("CWE Email Validator")

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

            # Validate each email and store only invalid ones
            invalid_emails = []
            for email in email_list:
                if is_valid_email_format(email):
                    domain = email.split('@')[1]
                    if not (has_valid_mx_record(domain) and smtp_check(email)):
                        invalid_emails.append(
                            (email, "Invalid MX record or SMTP check failed"))
                else:
                    invalid_emails.append((email, "Invalid format"))

            # Display the invalid emails
            if invalid_emails:
                st.write("Invalid Emails:")
                for email, reason in invalid_emails:
                    st.write(f"The email '{email}' is invalid: {reason}")
            else:
                st.write("All emails are valid.")
