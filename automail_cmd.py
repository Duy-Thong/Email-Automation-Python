import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import time

data_path = "data.xlsx"
template_path = "template.html"
sender_email = ""
sender_password = ""
subject = ""


def send_email(data_path, template_path, sender_email, sender_password, subject):
    # Determine the content type based on the file extension
    if template_path.endswith(".html"):
        content_type = "html"
    else:
        content_type = "plain"

    # Read the email content template
    with open(template_path, "r", encoding="utf-8") as template_file:
        email_content = template_file.read()

    # Read data from the Excel file
    wb = openpyxl.load_workbook(data_path)
    sheet = wb.active
    # Get the list of recipients from columns A (Full name) and B (Email)
    recipients = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        full_name, email, *_ = row  # Only consider columns A and B
        if email:
            recipients.append((full_name, email))

    # Connect to Gmail's SMTP server
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    try:
        # Log in to the email account
        server.login(sender_email, sender_password)
        # Send email to each recipient
        sent_successfully = []  # List of successfully sent emails
        failed_recipients = []  # List of failed emails

        for full_name, receiver_email in recipients:
            try:
                # Create a MIMEMultipart object to create the email
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = receiver_email
                message["Subject"] = subject
                email_content_with_name = email_content.replace("$NAME", full_name)
                # Attach the email content based on the determined content type
                message.attach(MIMEText(email_content_with_name, content_type))

                # Send the email
                server.sendmail(sender_email, receiver_email, message.as_string())

                # Record the successfully sent email
                sent_successfully.append((full_name, receiver_email))
                print(f"{full_name}, {receiver_email}")
            except Exception as e:
                # Record the failed email
                failed_recipients.append((full_name, receiver_email))
            time.sleep(1)

        # Close the SMTP connection
        server.quit()
    except Exception as e:
        print(f"Failed to send emails due to: {e}")
        # Close the SMTP connection in case of failure
        server.quit()
