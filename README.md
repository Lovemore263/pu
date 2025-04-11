Payslip Generator

Overview

This Python script automates the process of generating employee payslips in PDF format and emailing them to the respective employees. It reads employee data from an Excel file, generates a formatted payslip for each employee, and sends the payslip via email.

Features

Reads employee details from an Excel file.

Generates a professionally formatted PDF payslip.

Sends the payslip via email.

Uses environment variables for email credentials.

Handles errors and missing data gracefully.

Prerequisites

Ensure you have the following installed on your system:

Python 3.x

Required Python packages (install using pip):

pip install pandas fpdf python-dotenv smtplib

A valid email account (Gmail recommended) for sending payslips.

An Excel file containing employee details with the following required columns:

Employee ID

Name

Email

Basic Salary

Allowances

Deductions

Setup

Clone this repository or download the script.

Create a .env file in the same directory as the script and add the following environment variables:

EMAIL_ADDRESS=your-email@example.com
EMAIL_PASSWORD=your-email-password
SMTP_SERVER=smtp.gmail.com  # Change if using a different email provider
SMTP_PORT=587  # Default SMTP port for TLS

Place your Excel file (e.g., Employee details.xlsx) in the specified file path or modify the file_path variable in the script.

Usage

Run the script:

python payslip_generator.py

The script will:

Validate the presence of the Excel file and required columns.

Generate payslips and save them in the payslips/ directory.

Send the payslips to the corresponding employees via email.

Error Handling

If the .env file is missing or credentials are incorrect, the script will exit with an error.

If the Excel file is missing or has incorrect columns, an error message will be displayed.

If an employee's email is missing, their payslip will be skipped with a warning.

Any failures in email sending will be logged.

Notes

Ensure your email provider allows less secure apps or use an app-specific password.

You can customize the payslip format by modifying the generate_payslip() function.

Modify the email content in the send_email() function if needed.

License

This script is open-source and free to use for personal or business purposes.
