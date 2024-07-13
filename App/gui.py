import webbrowser
import customtkinter as ctk
import smtplib
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import os
import sys
import time
import concurrent.futures

# Global variables
sender_email = ""
sender_password = ""
data_path = ""
template_path = ""
subject = ""
success_path = "done.txt"
fail_path = "fail.txt"


def login():
    global sender_email, sender_password
    # Get sender email and password
    sender_email = email_entry.get()
    sender_password = password_entry.get()

    if sender_email == "" or sender_password == "":
        messagebox.showerror("Error", "Please enter your email and password.")
        return
    # Try to login with email and password
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    try:
        server.login(sender_email, sender_password)
        screen1.pack_forget()
        screen2.pack(fill="both", expand=True)
    except Exception as e:
        # Delete email and password
        password_entry.delete(0, ctk.END)
        messagebox.showerror(
            "Error", f"Login failed. Please check your email and password. {e}"
        )


def submitfile():
    global data_path, template_path
    data_path = data_path_entry.get()
    template_path = index_path_entry.get()
    if data_path == "" or template_path == "":
        messagebox.showerror("Empty Fields!", "Please enter data path and index path.")
        return
    screen2.pack_forget()
    screen3.pack(fill="both", expand=True)


def send_email():
    global subject
    subject = input_subject_entry.get()
    if subject == "":
        messagebox.showerror("Empty Field!", "Please enter the subject.")
        return
    send()
    messagebox.showinfo("Success", "Emails have been sent successfully!")


def back_to_screen1():
    screen2.pack_forget()
    screen1.pack(fill="both", expand=True)


def back_to_screen2():
    screen3.pack_forget()
    screen2.pack(fill="both", expand=True)


def user_manual():
    webbrowser.open(
        "https://docs.google.com/document/d/1I5GG2K_csg7bR1ZQf-5ldzEtqPN-6IJN7xn6MJDujWM/edit?usp=sharing"
    )


def send_email_task(recipient, content_type, subject, email_content):
    full_name, receiver_email = recipient
    try:
        # Connect to Gmail's SMTP server
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)

        # Create a MIMEMultipart object to create the email
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        email_content_with_name = email_content.replace("$NAME", full_name)
        message.attach(MIMEText(email_content_with_name, content_type))

        # Send the email
        server.sendmail(sender_email, receiver_email, message.as_string())

        # Close the SMTP connection
        server.quit()

        print(f"Email sent to {full_name} ({receiver_email})")
        return (full_name, receiver_email, True)
    except Exception as e:
        print(f"Failed to send email to {full_name} ({receiver_email}): {e}")
        return (full_name, receiver_email, False)


def send():
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

    # Use ThreadPoolExecutor to send emails concurrently
    sent_successfully = []  # List of successfully sent emails
    failed_recipients = []  # List of failed emails

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = [
            executor.submit(
                send_email_task, recipient, content_type, subject, email_content
            )
            for recipient in recipients
        ]
        for future in concurrent.futures.as_completed(futures):
            full_name, receiver_email, success = future.result()
            if success:
                sent_successfully.append((full_name, receiver_email))
            else:
                failed_recipients.append((full_name, receiver_email))

    with open(success_path, "w", encoding="utf-8") as f:
        for name, email in sent_successfully:
            f.write(f"{name}, {email}\n")

    with open(fail_path, "w", encoding="utf-8") as f:
        for name, email in failed_recipients:
            f.write(f"{name}, {email}\n")


app = ctk.CTk()
app.title("SoMedia Email Automation")
app.geometry("600x500")
app.resizable(False, False)

# Screen 1: input sender email and password
screen1 = ctk.CTkFrame(master=app)
screen1.pack(fill="both", expand=True)

# Set the background image (optional)
# script_directory = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
# image_path = os.path.join(script_directory, "assets/cover2.jpg")
# original_image = Image.open(image_path)
# resized_image = original_image.resize((800, 180))
# tk_image = ImageTk.PhotoImage(resized_image)
# canvas = ctk.CTkCanvas(master=screen1, width=800, height=200)
# canvas.create_image(0, 0, anchor=tk.NW, image=tk_image)
# canvas.pack()

text = ctk.CTkLabel(
    master=screen1,
    text="Welcome to Email Automation",
    font=("Arial", 20, "bold"),
    text_color="#402E7A",
)
text.pack(pady=30)

email_label = ctk.CTkLabel(master=screen1, text="Email", font=("Arial", 14))
email_label.pack()

email_entry = ctk.CTkEntry(master=screen1, width=300, height=30, font=("Arial", 12))
email_entry.pack(pady=5)

password_label = ctk.CTkLabel(master=screen1, text="Password", font=("Arial", 14))
password_label.pack()

password_entry = ctk.CTkEntry(
    master=screen1, show="*", width=300, height=30, font=("Arial", 12)
)
password_entry.pack(pady=5)

login_button = ctk.CTkButton(
    master=screen1,
    text="Login",
    width=100,
    height=40,
    command=login,
    corner_radius=10,
    fg_color="#4B70F5",
)
login_button.pack(pady=20)

user_manual_button = ctk.CTkButton(
    master=screen1,
    text="User Manual",
    width=100,
    height=40,
    command=user_manual,
    corner_radius=10,
    fg_color="#264653",
)
user_manual_button.pack()

# Screen 2: file selection
screen2 = ctk.CTkFrame(master=app)

back_button = ctk.CTkButton(
    master=screen2,
    text="Back",
    width=70,
    height=30,
    command=back_to_screen1,
    corner_radius=10,
    fg_color="#264653",
)
back_button.pack(side=tk.TOP, padx=20, pady=20)

data_path_label = ctk.CTkLabel(screen2, text="Path to xlsx file", font=("Arial", 14))
data_path_label.pack()

data_path_entry = ctk.CTkEntry(screen2, width=300, height=30, font=("Arial", 12))
data_path_entry.pack(pady=10)

select_data_button = ctk.CTkButton(
    screen2,
    text="Choose file",
    command=lambda: data_path_entry.insert(0, filedialog.askopenfilename()),
    width=100,
    height=40,
    corner_radius=10,
    fg_color="#2a9d8f",
)
select_data_button.pack()

index_path_label = ctk.CTkLabel(
    screen2, text="Path to content file (.html or .txt)", font=("Arial", 14)
)
index_path_label.pack()

index_path_entry = ctk.CTkEntry(screen2, width=300, height=30, font=("Arial", 12))
index_path_entry.pack(pady=10)

select_index_button = ctk.CTkButton(
    screen2,
    text="Choose file",
    command=lambda: index_path_entry.insert(0, filedialog.askopenfilename()),
    width=100,
    height=40,
    corner_radius=10,
    fg_color="#2a9d8f",
)
select_index_button.pack()

next_button = ctk.CTkButton(
    master=screen2,
    text="Next",
    width=100,
    height=40,
    command=submitfile,
    corner_radius=10,
    fg_color="#4B70F5",
)
next_button.pack(pady=20)

# Screen 3: input subject and send emails
screen3 = ctk.CTkFrame(master=app)

back_tosc2 = ctk.CTkButton(
    master=screen3,
    text="Back",
    width=70,
    height=30,
    command=back_to_screen2,
    corner_radius=10,
    fg_color="#264653",
)
back_tosc2.pack(side=tk.TOP, padx=20, pady=20)

input_subject_label = ctk.CTkLabel(screen3, text="Input subject:", font=("Arial", 14))
input_subject_label.pack()

input_subject_entry = ctk.CTkEntry(screen3, width=300, height=30, font=("Arial", 12))
input_subject_entry.pack(pady=10)

send_button = ctk.CTkButton(
    screen3,
    text="Send",
    width=100,
    height=40,
    command=send_email,
    corner_radius=10,
    fg_color="#4B70F5",
)
send_button.pack(pady=20)

done_button = ctk.CTkButton(
    master=screen3,
    text="Done",
    width=100,
    height=40,
    command=app.quit,
    corner_radius=10,
    fg_color="#264653",
)
done_button.pack(pady=20)

to_success = ctk.CTkLabel(
    master=screen3,
    text="Successfully sent emails are recorded in done.txt",
    font=("Arial", 15),
    text_color="#2a9d8f",
)
to_success.pack(pady=50)

app.mainloop()
