import os
import pandas as pd
from fpdf import FPDF
from email.message import EmailMessage
import smtplib
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

# Debug check for email credentials
if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
    print("❌ ERROR: EMAIL_ADDRESS or EMAIL_PASSWORD not set in .env")
    exit()

# Create payslips folder if it doesn't exist
os.makedirs("payslips", exist_ok=True)

# Load Excel file
file_path = r"C:\Users\uncommonstudent\Downloads\Employee details.xlsx"

if not os.path.exists(file_path):
    print("❌ ERROR: Excel file not found.")
    exit()
else:
    print(f"✅ Found Excel file at {file_path}")

# Read employee data
try:
    df = pd.read_excel(file_path, dtype={"Employee ID": str})  # Ensure Employee ID is string
    df.columns = df.columns.str.strip()  # Clean column headers
except Exception as e:
    print(f"❌ ERROR: Unable to read Excel file: {e}")
    exit()

# Check required columns
required_columns = {"Employee ID", "Name", "Email", "Basic Salary", "Allowances", "Deductions"}
if not required_columns.issubset(set(df.columns)):
    print(f"❌ ERROR: Missing required columns in Excel file. Expected: {required_columns}")
    exit()

# Convert numeric columns to float
df["Basic Salary"] = pd.to_numeric(df["Basic Salary"], errors="coerce").fillna(0)
df["Allowances"] = pd.to_numeric(df["Allowances"], errors="coerce").fillna(0)
df["Deductions"] = pd.to_numeric(df["Deductions"], errors="coerce").fillna(0)

# Payslip PDF Generator
def generate_payslip(emp):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Calculate net salary
        net_salary = emp["Basic Salary"] + emp["Allowances"] - emp["Deductions"]

        # Payslip content
        lines = [
            f"Employee ID: {emp['Employee ID']}",
            f"Name: {emp['Name']}",
            f"Basic Salary: ${emp['Basic Salary']:.2f}",
            f"Allowances: ${emp['Allowances']:.2f}",
            f"Deductions: ${emp['Deductions']:.2f}",
            f"Net Salary: ${net_salary:.2f}",
        ]

        for line in lines:
            pdf.cell(200, 10, txt=line, ln=True)

        filename = f"payslips/{emp['Employee ID']}.pdf"
        pdf.output(filename)
        return filename
    except Exception as e:
        print(f"❌ ERROR: Failed to generate payslip for {emp['Name']}: {e}")
        return None

# Email Sender Function
def send_email(to_email, filename, name):
    if not to_email or pd.isna(to_email):
        print(f"⚠️ WARNING: No email provided for {name}. Skipping...")
        return False

    msg = EmailMessage()
    msg["Subject"] = "Your Payslip for This Month"
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg.set_content(f"Dear {name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR")

    # Attach payslip PDF
    try:
        with open(filename, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=os.path.basename(filename))

        # Send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        
        return True
    except Exception as e:
        print(f"❌ ERROR: Failed to send email to {name} ({to_email}): {e}")
        return False

# Process each employee
for _, emp in df.iterrows():
    try:
        payslip_file = generate_payslip(emp)
        if payslip_file:
            sent = send_email(emp["Email"], payslip_file, emp["Name"])
            if sent:
                print(f"✅ Payslip sent to {emp['Name']} ({emp['Email']})")
    except Exception as e:
        print(f"❌ ERROR: Processing failed for {emp['Name']} ({emp['Email']}): {e}")

