import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl

# Paths to the files
data_path = r'D:\Work_space\Email-Automation-Python\Infor.xlsx'   # Replace with the path to your Excel file
index_path = r'D:\Work_space\Email-Automation-Python\index.html'

# Your email account information
sender_email = "YOUR_EMAIL_ADDRESS"
sender_password = "YOUR_EMAIL_PASSWORD"

# Read email content from the HTML file
with open(index_path, "r", encoding="utf-8") as html_file:
    email_content = html_file.read()

# Read data from the Excel file
wb = openpyxl.load_workbook(data_path)
sheet = wb.active

# Get a list of recipient information from column A (Full name) and column B (Email)
recipients = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    full_name, email, *_ = row  # Only interested in columns A and B
    if email:
        recipients.append((full_name, email))

# Connect to the Gmail SMTP server
smtp_server = "smtp.gmail.com"
smtp_port = 587
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()

# Log in to the email account
server.login(sender_email, sender_password)

# Send emails to each recipient
for full_name, receiver_email in recipients:
    # Create a MIMEMultipart object to compose the email
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "[ NOTIFICATION OF ROUND CV RESULTS - MEDIA BOOK ]"
    email_content_with_name = email_content.replace("$NAME", full_name)
    # Email content is taken from the HTML file
    message.attach(MIMEText(email_content_with_name, "html"))

    # Send the email
    server.sendmail(sender_email, receiver_email, message.as_string())

# Close the SMTP connection
server.quit()

print("Emails have been sent successfully!")
