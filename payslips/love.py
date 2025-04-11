import os
import pandas as pd
from fpdf import FPDF
from dotenv import load_dotenv
import time  # Import the time module to use sleep()

# Load environment variables
load_dotenv()
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

# Create payslips folder if it doesn't exist
os.makedirs("payslips", exist_ok=True)

# Load Excel file
file_path = r"C:\Users\uncommonstudent\Downloads\Employee details.xlsx"
if not os.path.exists(file_path):
    print("❌ ERROR: Excel file not found.")
    exit()
print(f"✅ Found Excel file at {file_path}")

# Read employee data
df = pd.read_excel(file_path, dtype={"Employee ID": str})
df.columns = df.columns.str.strip()

df["Basic Salary"] = pd.to_numeric(df["Basic Salary"], errors="coerce").fillna(0)
df["Allowances"] = pd.to_numeric(df["Allowances"], errors="coerce").fillna(0)
df["Deductions"] = pd.to_numeric(df["Deductions"], errors="coerce").fillna(0)

# Payslip PDF Generator
class Payslip(FPDF):
    def header(self):
        self.set_font("Arial", "B", 16)
        self.set_text_color(0, 0, 128)  # Dark Blue
        self.cell(200, 10, "STEADYFINGERS PAYSLIP", ln=True, align="C")
        self.ln(5)
        self.set_draw_color(0, 0, 128)
        self.line(10, 20, 200, 20)  # Top border

    def payslip_section(self, title, value, color=(0, 0, 0)):
        self.set_font("Arial", "B", 12)
        self.set_text_color(*color)
        self.cell(100, 10, title, border=1)
        self.set_font("Arial", "", 12)
        self.cell(90, 10, value, border=1, ln=True)


def generate_payslip(emp):
    pdf = Payslip()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.ln(10)

    # Employee details
    pdf.payslip_section("Employee ID:", emp["Employee ID"])
    pdf.payslip_section("Name:", emp["Name"])
    pdf.payslip_section("Basic Salary:", f"${emp['Basic Salary']:.2f}", (0, 128, 0))
    pdf.payslip_section("Allowances:", f"${emp['Allowances']:.2f}", (0, 128, 0))
    pdf.payslip_section("Deductions:", f"-${emp['Deductions']:.2f}", (255, 0, 0))
    
    # Calculate net salary
    net_salary = emp["Basic Salary"] + emp["Allowances"] - emp["Deductions"]
    pdf.payslip_section("Net Salary:", f"${net_salary:.2f}", (0, 0, 255))
    
    filename = f"payslips/{emp['Employee ID']}.pdf"
    pdf.output(filename)
    return filename


# Process each employee
for _, emp in df.iterrows():
    payslip_file = generate_payslip(emp)
    print(f"✅ Payslip generated for {emp['Name']}")
    
    # Add a delay (e.g., 1 second between each payslip generation)
    time.sleep(1)  # This will pause for 1 second before moving to the next employee
