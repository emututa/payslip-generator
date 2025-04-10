import pandas as pd
from fpdf import FPDF
import yagmail
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Read Excel
df = pd.read_excel("employees.xlsx")

# Create output folder
os.makedirs("payslips", exist_ok=True)

# Initialize email client
yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASSWORD)

# Generate and send payslips
for index, row in df.iterrows():
    try:
        emp_id = row['Employee ID']
        name = row['Name']
        email = row['Email']
        basic = row['Basic Salary']
        allowance = row['Allowances']
        deduction = row['Deductions']
        net_salary = basic + allowance - deduction

        # Generate PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt=f"Payslip for {name} (ID: {emp_id})", ln=True, align='C')
        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Basic Salary: ${basic}", ln=True)
        pdf.cell(200, 10, txt=f"Allowances: ${allowance}", ln=True)
        pdf.cell(200, 10, txt=f"Deductions: ${deduction}", ln=True)
        pdf.cell(200, 10, txt=f"Net Salary: ${net_salary}", ln=True)
        file_path = f"payslips/{emp_id}.pdf"
        pdf.output(file_path)

        # Send email
        yag.send(
            to=email,
            subject="Your Payslip for This Month",
            contents=f"Hi {name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nDDIS3609006",
            attachments=file_path
        )
        print(f"Sent payslip to {name} ({email})")
    except Exception as e:
        print(f"Error processing {row['Name']}: {e}")
