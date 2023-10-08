import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import tkinter as tk
from tkinter import filedialog
import time

# Hàm để gửi email


def send_email():
    # Lấy thông tin từ các trường nhập liệu
    data_path = data_path_entry.get()
    index_path = index_path_entry.get()
    sender_email = email_entry.get()
    sender_password = password_entry.get()
    subject = subject_entry.get()

    # Đọc nội dung email từ tệp HTML
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
    smtp_port = 4
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

        # Ghi danh sách người nhận đã được gửi email thành công vào file done.txt
        with open("done.txt", "w") as done_file:
            for name, email in sent_successfully:
                done_file.write(f"{email}\n")

        # Ghi danh sách email gửi thất bại vào file notdone.txt
        with open("notdone.txt", "w") as notdone_file:
            for name, email in failed_recipients:
                notdone_file.write(f"{email}\n")

        result_label.config(text="Email đã được gửi thành công!")
    except smtplib.SMTPAuthenticationError:
        result_label.config(
            text="Lỗi: Đăng nhập vào tài khoản email thất bại.")


# Tạo giao diện đồ họa
root = tk.Tk()
root.title("Gửi Email")

# Label và Entry cho đường dẫn đến file Excel
data_path_label = tk.Label(root, text="Đường dẫn đến file Excel:")
data_path_label.pack()
data_path_entry = tk.Entry(root)
data_path_entry.pack()

# Button để chọn file Excel
select_data_button = tk.Button(root, text="Chọn file Excel",
                               command=lambda: data_path_entry.insert(0, filedialog.askopenfilename()))
select_data_button.pack()

# Label và Entry cho đường dẫn đến file HTML
index_path_label = tk.Label(root, text="Đường dẫn đến file HTML:")
index_path_label.pack()
index_path_entry = tk.Entry(root)
index_path_entry.pack()

# Button để chọn file HTML
select_index_button = tk.Button(
    root, text="Chọn file HTML", command=lambda: index_path_entry.insert(0, filedialog.askopenfilename()))
select_index_button.pack()

# Label và Entry cho địa chỉ email và mật khẩu
email_label = tk.Label(root, text="Email:")
email_label.pack()
email_entry = tk.Entry(root)
email_entry.pack()

password_label = tk.Label(root, text="Mật khẩu:")
password_label.pack()
password_entry = tk.Entry(root, show="*")
password_entry.pack()

# Label và Entry cho tiêu đề email
subject_label = tk.Label(root, text="Tiêu đề email:")
subject_label.pack()
subject_entry = tk.Entry(root)
subject_entry.pack()

# Button để gửi email
send_button = tk.Button(root, text="Gửi Email", command=send_email)
send_button.pack()

# Label để hiển thị kết quả
result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
