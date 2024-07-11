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
        screen2.pack()
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
    screen3.pack()


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
    screen1.pack()


def back_to_screen2():
    screen3.pack_forget()
    screen2.pack()


def user_manual():
    webbrowser.open(
        "https://docs.google.com/document/d/1I5GG2K_csg7bR1ZQf-5ldzEtqPN-6IJN7xn6MJDujWM/edit?usp=sharing"
    )


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
                print(f"Failed to send email to {full_name} ({receiver_email}): {e}")
            time.sleep(1)
        with open(success_path, "w", encoding="utf-8") as f:
            for name, email in sent_successfully:
                f.write(f"{name}, {email}\n")

        with open(fail_path, "w", encoding="utf-8") as f:
            for name, email in failed_recipients:
                f.write(f"{name}, {email}\n")
        # Close the SMTP connection
        server.quit()

    except Exception as e:
        print(f"Failed to send emails due to: {e}")
        # Close the SMTP connection in case of failure
        server.quit()


app = ctk.CTk()
app.title("SoMedia Email Automation")
app.geometry("600x400")
app.resizable(False, False)

# Screen 1: input sender email and password
screen1 = ctk.CTkFrame(master=app, width=800, height=600)
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
    font=("Arial", 15),
    text_color="black",
)
text.pack(pady=50)

email_label = ctk.CTkLabel(master=screen1, text="Email")
email_label.pack()

email_entry = ctk.CTkEntry(master=screen1, width=200, height=30)
email_entry.pack()

password_label = ctk.CTkLabel(master=screen1, text="Password")
password_label.pack()

password_entry = ctk.CTkEntry(master=screen1, show="*", width=200, height=30)
password_entry.pack()

login_button = ctk.CTkButton(
    master=screen1, text="Login", width=100, height=30, command=login
)
login_button.pack(pady=20)

user_manual_button = ctk.CTkButton(
    master=screen1, text="User Manual", width=100, height=30, command=user_manual
)
user_manual_button.pack()

# ___________________________________________________________________
# Screen 2:
screen2 = ctk.CTkFrame(master=app, width=800, height=600)

back_button = ctk.CTkButton(
    master=screen2, text="Back", width=70, height=25, command=back_to_screen1
)
back_button.pack(side=tk.TOP, padx=20, pady=20)

data_path_label = ctk.CTkLabel(screen2, text="Path to xlsx file")
data_path_label.pack()

data_path_entry = ctk.CTkEntry(screen2)
data_path_entry.pack(pady=10)

select_data_button = ctk.CTkButton(
    screen2,
    text="Choose file",
    command=lambda: data_path_entry.insert(0, filedialog.askopenfilename()),
)
select_data_button.pack()

index_path_label = ctk.CTkLabel(screen2, text="Path to content file (.html or .txt)")
index_path_label.pack()

index_path_entry = ctk.CTkEntry(screen2)
index_path_entry.pack(pady=10)

select_index_button = ctk.CTkButton(
    screen2,
    text="Choose file",
    command=lambda: index_path_entry.insert(0, filedialog.askopenfilename()),
)
select_index_button.pack()

Next_button = ctk.CTkButton(
    master=screen2, text="Next", width=100, height=30, command=submitfile
)
Next_button.pack(pady=20)

# ____________________________________________________________________#
# Screen 3:

screen3 = ctk.CTkFrame(master=app, width=400, height=300)

back_tosc2 = ctk.CTkButton(
    master=screen3, text="Back", width=70, height=30, command=back_to_screen2
)
back_tosc2.pack(side=tk.TOP, padx=20, pady=20)

input_subject_label = ctk.CTkLabel(screen3, text="Input subject:")
input_subject_label.pack()

input_subject_entry = ctk.CTkEntry(screen3, width=200, height=30)
input_subject_entry.pack()

send_button = ctk.CTkButton(
    screen3, text="Send", width=100, height=30, command=send_email
)
send_button.pack(pady=20)
done_button = ctk.CTkButton(
    master=screen3, text="Done", width=100, height=30, command=app.quit
)
done_button.pack(pady=20)
to_success = ctk.CTkLabel(
    master=screen3,
    text="Successfully sent emails are recorded in done.txt",
    font=("Arial", 15),
    text_color="black",
)
to_success.pack(pady=50)

# _______________________________________________________#

app.mainloop()
