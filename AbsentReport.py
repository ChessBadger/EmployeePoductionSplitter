import PyPDF2
from fpdf import FPDF

# Function to extract data from the provided PDF
def extract_data_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = "".join(page.extract_text() for page in reader.pages)
    return text

# Function to parse the text and extract employee data
def parse_employee_data(text):
    employees = []
    lines = text.splitlines()
    for line in lines:
        if '-' in line and 'Value' in line:
            parts = line.split('-')
            if len(parts) > 1:
                name = parts[0].strip()
                points = int(parts[1].split('Value')[0].strip())
                employees.append((name, points))
    return employees

# Function to create a sorted PDF
def create_sorted_pdf(employees, output_path):
    employees_sorted = sorted(employees, key=lambda x: x[1], reverse=True)

    pdf = FPDF()
    pdf.set_font("Arial", size=12)
    pdf.add_page()

    pdf.cell(200, 10, txt="Employee Points Report", ln=True, align='C')
    pdf.ln(10)

    for name, points in employees_sorted:
        pdf.cell(200, 10, txt=f"{name}: {points} points", ln=True)

    pdf.output(output_path)

# Paths
input_pdf_path = r'C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\CallInsReport.pdf'
output_pdf_path = r'C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\SortedEmployeePoints.pdf'

# Execution
pdf_text = extract_data_from_pdf(input_pdf_path)
employee_data = parse_employee_data(pdf_text)
create_sorted_pdf(employee_data, output_pdf_path)

print(f"Sorted PDF created at {output_pdf_path}")
