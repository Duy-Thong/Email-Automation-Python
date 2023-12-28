import webbrowser
import customtkinter as ctk
import smtplib
import tkinter as tk
from tkinter import filedialog
from PIL import Image,ImageTk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import os
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import tkinter as tk
from tkinter import filedialog
import time
# Global variables
sender_email = ""
sender_password = ""
data_path = ""
index_path = ""
subject = ""

def login():
    global sender_email, sender_password
    # Get sender email and password
    sender_email = email_entry.get()
    sender_password = password_entry.get()
    
    if sender_email == "" or sender_password == "":
        tk.messagebox.showerror("Error", "Please enter your email and password.")
        return
    # Try to login with email and password
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    try:
        server.login(sender_email, sender_password)
        screen1.pack_forget()
        screen2.pack()
    except:
        # Delete email and password
        password_entry.delete(0, ctk.END)
        tk.messagebox.showerror("Error", "Login failed. Please check your email and password.")

def submitfile():
    global data_path, index_path
    data_path = data_path_entry.get()
    index_path = index_path_entry.get()
    if data_path == "" or index_path == "":
        tk.messagebox.showerror("Empty Fields!", "Please enter data path and index path.")
        return
    screen2.pack_forget()
    screen3.pack()

def send_email():
    global subject
    subject = input_subject_entry.get()
    if "html" in index_path:
        send_html_email()
    elif "txt" in index_path:
        send_text_email()
def send_html_email():
    send_email_html()
    tk.messagebox.showinfo("Success", "Email sent successfully!")
def send_text_email():
    send_email_txt()
    tk.messagebox.showinfo("Success", "Email sent successfully!")

def back_to_screen1():
    screen2.pack_forget()
    screen1.pack()

def back_to_screen2():
    screen3.pack_forget()
    screen2.pack()
def user_manual():
    webbrowser.open("https://docs.google.com/document/d/1I5GG2K_csg7bR1ZQf-5ldzEtqPN-6IJN7xn6MJDujWM/edit?usp=sharing")

def send_email_html():
    with open(index_path, "r", encoding="utf-8") as html_file:
        email_content = html_file.read()

    # Đọc dữ liệu từ file Excel
    wb = openpyxl.load_workbook(data_path)
    sheet = wb.active
    # Lấy danh sách thông tin người nhận từ cột A (Full name) và cột B (Email)
    recipients = []
    cnt=1
    for row in sheet.iter_rows(min_row=2, values_only=True):
        cnt+=1
        full_name, email, *_ = row  # Chỉ quan tâm đến cột A và cột B
        if email:
            recipients.append((full_name, email))
            sheet.cell(row=cnt, column=3).value = True
    wb.save(data_path)
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
def send_email_txt():

    with open(index_path, "r", encoding="utf-8") as text_file:
        email_content = text_file.read()
    # Đọc dữ liệu từ file Excel
    wb = openpyxl.load_workbook(data_path)
    sheet = wb.active
    # Lấy danh sách thông tin người nhận từ cột A (Full name) và cột B (Email)
    recipients = []
    cnt=1
    for row in sheet.iter_rows(min_row=2, values_only=True):
        full_name, email, *_ = row  # Chỉ quan tâm đến cột A và cột B
        cnt+=1
        if email:
            recipients.append((full_name, email))
            sheet.cell(row=cnt, column=3).value = True
    wb.save(data_path)
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
              

app = ctk.CTk()
app.title("Email Automation by DuyThong")
app.geometry("600x400")
app.resizable(False, False)

# Screen 1: input sender email and password
screen1 = ctk.CTkFrame(master=app, width=800, height=600)
screen1.pack(fill="both", expand=True)

script_directory = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

# image_path = r"Email-Automation-Python/assets/cover2.jpg"
# original_image = Image.open(image_path)
# resized_image = original_image.resize((800,160))
# tk_image = ImageTk.PhotoImage(resized_image)

# Set the background image
canvas = ctk.CTkCanvas(master=screen1, width=800, height=160,bg="cyan")
canvas.pack()

text = ctk.CTkLabel(master=screen1, text="Welcome to Email Automation by DuyThong", font=("Arial", 15), text_color="white",text_color_disabled="white")
text.place(x=150,y=100)
# Create fields to input email and password
email_label = ctk.CTkLabel(master=screen1, text="Email")
email_label.place(x=300, y=100)
email_label.pack()

email_entry = ctk.CTkEntry(master=screen1,width=200, height=30)
email_entry.place(x=300, y=140)
email_entry.pack()

password_label = ctk.CTkLabel(master=screen1, text="Password")
password_label.place(x=300, y=180)
password_label.pack()

password_entry = ctk.CTkEntry(master=screen1, show="*",width=200, height=30)
password_entry.place(x=300, y=220)
password_entry.pack()

login_button = ctk.CTkButton(master=screen1, text="Login", width=100, height=30, command=lambda: login())
login_button.place(relx=0.5, rely=0.8, anchor=tk.CENTER)
login_button.pack()


user_manual_button = ctk.CTkButton(master=screen1, text="User Manual", width=100, height=30, command=lambda: user_manual())
user_manual_button.place(relx=0.5, rely=0.9, anchor=tk.CENTER)
user_manual_button.pack()
#___________________________________________________________________
# Screen 2:
screen2 = ctk.CTkFrame(master=app, width=800, height=600, bg_color="blue")


# Set the background image
canvas = ctk.CTkCanvas(master=screen2, width=800, height=160,bg="cyan")
canvas.pack()


back_button = ctk.CTkButton(master=screen2, text="Back", width=70, height=25, command=lambda: back_to_screen1())
back_button.pack()
back_button.place(x=10, y=10)
data_path_label = ctk.CTkLabel(screen2, text="Path to xlsx file")
data_path_label.pack()
data_path_entry = ctk.CTkEntry(screen2)
data_path_entry.pack()

select_data_button = ctk.CTkButton(screen2, text="Choose file",command=lambda: data_path_entry.insert(0, filedialog.askopenfilename()))
select_data_button.pack()

index_path_label = ctk.CTkLabel(screen2, text="Path to content file (.html or .txt)")
index_path_label.pack()
index_path_entry = ctk.CTkEntry(screen2)
index_path_entry.pack()

# CTkButton để chọn file HTML
select_index_button = ctk.CTkButton(screen2, text="Choose file", command=lambda: index_path_entry.insert(0, filedialog.askopenfilename()))
select_index_button.pack()

Next_button = ctk.CTkButton(master=screen2, text="Next", width=100, height=30, command=lambda: submitfile())
Next_button.pack()
#____________________________________________________________________#
#Screen 3:

screen3=ctk.CTkFrame(master=app, width=400, height=300, bg_color="blue")


# Set the background image
canvas = ctk.CTkCanvas(master=screen3, width=800, height=160,bg="cyan")
canvas.pack()

back_tosc2=ctk.CTkButton(master=screen3, text="Back", width=70, height=30, command=lambda: back_to_screen2())
back_tosc2.place(relx=0.5, rely=0.9)
back_tosc2.pack()
input_subject_label = ctk.CTkLabel(screen3, text="Input subject:")
input_subject_label.pack()

input_subject_entry = ctk.CTkEntry(screen3,width=200, height=30)
input_subject_entry.pack()

send_button = ctk.CTkButton(screen3, text="Send",width=100, height=30, command=lambda: send_email())
send_button.place(x=300, y=260)
send_button.pack()
#_______________________________________________________#
app.mainloop()
