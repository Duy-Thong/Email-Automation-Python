import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
#Đường dẫn đến các file 
data_path=r'D:\Work_space\Email-Automation-Python\Infor.xlsx'   #thay đường dẫn bằng đường dẫn đến file xlxs
index_path=r'D:\Work_space\Email-Automation-Python\index.html'

# Thông tin về tài khoản email của bạn
sender_email = "clbsomediaptit@gmail.com"
sender_password = "gvszzolbdxqlviyi"

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
smtp_port = 587
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()

# Đăng nhập vào tài khoản email
server.login(sender_email, sender_password)

# Gửi email cho từng người nhận
for full_name, receiver_email in recipients:
    # Tạo đối tượng MIMEMultipart để tạo email
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "[ THÔNG BÁO KẾT QUẢ VÒNG CV - SỔ MEDIA ]"
    email_content_with_name = email_content.replace("$NAME", full_name)
    # Nội dung email lấy từ tệp HTML
    message.attach(MIMEText(email_content_with_name, "html"))

    # Gửi email
    server.sendmail(sender_email, receiver_email, message.as_string())

# Đóng kết nối SMTP
server.quit()

print("Email đã được gửi thành công!")
