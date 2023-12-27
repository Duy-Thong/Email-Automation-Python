import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import tkinter as tk
from tkinter import filedialog
import time

def send_email(data_path, index_path, sender_email, sender_password,subject):
    

    with open(index_path, "r", encoding="utf-8") as html_file:
        email_content = html_file.read()

    # Đọc dữ liệu từ file Excel
    wb = openpyxl.load_workbook(data_path)
    sheet = wb.active
    # Lấy danh sách thông tin người nhận từ cột A (Full name) và cột B (Email)
    recipients = []
    for row in sheet.iter_rows(min_row=2, values_only=True):

        full_name, email, *_ = row  # Chỉ quan tâm đến cột A và cột B
        if email:
            recipients.append((full_name, email))
    
    # Kết nối đến máy chủ SMTP của Gmail
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    try:
        # Đăng nhập vào tài khoản email
        server.login(sender_email, sender_password)
        # Gửi email cho từng người nhận
        sent_successfully = []  # Danh sách người nhận đã được gửi email thành công
        failed_recipients = []  # Danh sách email gửi thất bại

        for full_name, receiver_email in recipients:
            try:
                # Tạo đối tượng MIMEMultipart để tạo email
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = receiver_email
                message["Subject"] = subject
                email_content_with_name = email_content.replace(
                    "$NAME", full_name)
                # Nội dung email lấy từ tệp HTML
                message.attach(MIMEText(email_content_with_name, "html"))

                # Gửi email
                server.sendmail(sender_email, receiver_email,
                                message.as_string())

                # Ghi nhận người nhận đã được gửi email thành công
                sent_successfully.append((full_name, receiver_email))
                print(f"{full_name} , {receiver_email}")
            except Exception as e:
                # Ghi nhận email gửi thất bại
                failed_recipients.append((full_name, receiver_email))
            time.sleep(1)

        # Đóng kết nối SMTP
        server.quit()
    except:
        # Đóng kết nối SMTP
        server.quit()
def send_email_txt(data_path, index_path, sender_email, sender_password,subject):

    with open(index_path, "r", encoding="utf-8") as text_file:
        email_content = text_file.read()
    # Đọc dữ liệu từ file Excel
    wb = openpyxl.load_workbook(data_path)
    sheet = wb.active
    # Lấy danh sách thông tin người nhận từ cột A (Full name) và cột B (Email)
    recipients = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        full_name, email, *_ = row  # Chỉ quan tâm đến cột A và cột B
        if email:
            recipients.append((full_name, email))

    # Kết nối đến máy chủ SMTP của Gmail
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    try:
        # Đăng nhập vào tài khoản email
        server.login(sender_email, sender_password)
        # Gửi email cho từng người nhận
        sent_successfully = []  # Danh sách người nhận đã được gửi email thành công
        failed_recipients = []  # Danh sách email gửi thất bại

        for full_name, receiver_email in recipients:
            try:
                # Tạo đối tượng MIMEMultipart để tạo email
                message = MIMEMultipart()
                message["From"] = sender_email
                message["To"] = receiver_email
                message["Subject"] = subject
                email_content_with_name = email_content.replace(
                    "$NAME", full_name)
                # Nội dung email lấy từ tệp HTML
                message.attach(MIMEText(email_content_with_name, "plain"))

                # Gửi email
                server.sendmail(sender_email, receiver_email,
                                message.as_string())

                # Ghi nhận người nhận đã được gửi email thành công
                sent_successfully.append((full_name, receiver_email))
                print(f"{full_name} , {receiver_email}")
            except Exception as e:
                # Ghi nhận email gửi thất bại
                failed_recipients.append((full_name, receiver_email))
            time.sleep(1)
        # Đóng kết nối SMTP
        server.quit()
    except:
        # Đóng kết nối SMTP
        server.quit()
              

