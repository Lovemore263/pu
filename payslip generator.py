import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import yagmail

# Load environment variables (set these before running the script)
EMAIL_SENDER = os.getenv("EMAIL_SENDER")  # Your email
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")  # Your email password
SMTP_SERVER = "smtp.gmail.com"  # Change this for other providers
SMTP_PORT = 587

# Read employee data
def read_employee_data(file_path):
    df = pd.read_excel(file_path)
    return df

# Generate payslip PDF
def create_payslip(employee):
    emp_id = employee["Employee ID"]
    name = employee["Name"]
    email = employee["Email"]
    basic_salary = employee["Basic Salary"]
    allowances = employee["Allowances"]
    deductions = employee["Deductions"]
    net_salary = basic_salary + allowances - deductions

    # Define file path
    payslip_dir = "payslips"
    os.makedirs(payslip_dir, exist_ok=True)  # Create directory if not exists
    file_path = os.path.join(payslip_dir, f"{emp_id}.pdf")

    # Create PDF
    c = canvas.Canvas(file_path, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(200, 800, "Company Payslip")
    c.setFont("Helvetica", 12)
    c.drawString(100, 750, f"Employee ID: {emp_id}")
    c.drawString(100, 730, f"Name: {name}")
    c.drawString(100, 710, f"Basic Salary: ${basic_salary:.2f}")
    c.drawString(100, 690, f"Allowances: ${allowances:.2f}")
    c.drawString(100, 670, f"Deductions: ${deductions:.2f}")
    c.drawString(100, 650, f"Net Salary: ${net_salary:.2f}")
    c.save()

    return file_path  # Return file path for emailing

# Send email with payslip attachment
def send_email(employee, payslip_path):
    recipient_email = employee["Email"]
    subject = "Your Payslip for This Month"
    body = f"Hello {employee['Name']},\n\nPlease find attached your payslip for this month.\n\nBest Regards,\nHR Department"

    try:
        yag = yagmail.SMTP(EMAIL_SENDER, EMAIL_PASSWORD)
        yag.send(to=recipient_email, subject=subject, contents=body, attachments=payslip_path)
        print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Error sending email to {recipient_email}: {e}")

# Main function to process employees
def process_payslips(file_path):
    employees = read_employee_data(file_path)

    for _, employee in employees.iterrows():
        payslip_path = create_payslip(employee)
        send_email(employee, payslip_path)

# Run the script
if __name__ == "__main__":
    process_payslips("employees.xlsx")
