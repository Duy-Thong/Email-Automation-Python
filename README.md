# Email Automation Tool

This Python application utilizes tkinter for the graphical interface and SMTP for email sending. It allows users to send personalized emails to a specified list of recipients from an Excel sheet.

## Overview

Instead of manually looping through and sending emails to a group of people, which can be inconvenient or impersonal, this tool sends batch emails individually, personalized with recipient names. Currently, it supports replacing names using the $NAME variable in the email content. Future updates may include additional variables for positions or addresses.

## Usage Guide

### Login to Email Account

#### Login Screen
Enter your email and password. Note: Use an App Password if 2-step verification is enabled for your Google account.

##### How to create an App Password:
- Follow Google's instructions: [Google Help Center](https://support.google.com/accounts/answer/185833)
- Ensure 2-step verification is enabled: Go to https://myaccount.google.com/security, select "Security" tab, turn on 2-step verification, and navigate to "App Passwords". Generate a new App Password and use it for login.

### Input Recipient List and Email Content

#### File Selection Screen
- **Excel File:** Should contain recipient details in the format: Column A: Full Name, Column B: Email. First row as header or leave empty.
- **Content File:** Either HTML or plain text format. Use $NAME where the recipient's name should be inserted.
  
##### Example:
- **HTML:** `<p>Hello, $NAME</p>`
- **TXT:** `Hello, $NAME`

### Enter Email Subject

#### Subject Input Screen
Enter the email subject and click "Send" to initiate email sending process. The app will automatically send emails to the recipients listed in the Excel sheet.

#### Note:
- After sending, a success message will display.
- Successful email deliveries will be marked True in the Excel sheet.


## Credits

Developed by Đào Duy Thông - D21 PTIT

GitHub Repository: [Email Automation Python](https://github.com/Duy-Thong/Email-Automation-Python)

