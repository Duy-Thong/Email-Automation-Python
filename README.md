
# Email Automation Python Script

This Python script automates the process of sending emails to a list of recipients using data from an Excel file and an HTML email template.

## Requirements

- Python 3.x
- `smtplib` library
- `email` library
- `openpyxl` library

## Setup

1. Clone or download this repository to your local machine.

2. Install the required Python libraries using pip:

   ```shell
   pip install openpyxl
   ```

3. Update the script with your email and server details:
   
   - Set the `data_path` variable to the path of your Excel file containing recipient data.
   - Set the `index_path` variable to the path of your HTML email template.
   - Replace the `sender_email` and `sender_password` variables with your email credentials.
   
4. Customize the email content in the HTML template (`index.html`). You can use the `$NAME` placeholder to personalize the email content with the recipient's name.

## Usage

1. Run the script using the following command:

   ```shell
   python send_email.py
   ```

2. The script will read the recipient data from the Excel file, personalize the email content, and send emails to the recipients with a throttling rate of 20 emails every 3 minutes to avoid being flagged as spam.

3. After the script completes, you will see a success message indicating that the emails have been sent.
4. Remember to replace your specifics like email address, password, html file address and xlsx file address with your specifics.

## Author

DAO DUY THONG

If you have any questions or issues, feel free to contact me at duythong.ptit@gmail.com.


