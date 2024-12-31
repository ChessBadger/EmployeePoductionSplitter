import re
import PyPDF2
from fpdf import FPDF

# Function to extract data from the provided PDF and filter out unwanted substrings
def extract_data_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = "".join(page.extract_text() for page in reader.pages)
    
    # Remove substrings matching the date format and page numbering format
    text = re.sub(r"\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday), [A-Za-z]+ \d{1,2}, \d{4}\b", "", text)
    text = re.sub(r"Page \d+ of \d+", "", text)
    return text

# Function to parse the text and extract employee data
def parse_employee_data(text):
    employees = []
    lines = text.splitlines()
    for line in lines:
        if '-' in line and 'Value' in line:
            parts = line.split(' - ')
            if len(parts) > 1:
                name = parts[0].strip()
                try:
                    points = int(parts[1].split('Value')[0].strip())
                    if points > 0:  # Exclude employees with zero points
                        employees.append((name, points))
                except ValueError:
                    print(f"Skipping line due to invalid points format: {line}")
    return employees

# Function to create a sorted PDF
def create_sorted_pdf(employees, output_path):
    employees_sorted = sorted(employees, key=lambda x: x[1], reverse=True)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Title
    pdf.set_font("Arial", style="B", size=16)
    pdf.cell(200, 10, txt="Employee Points Report", ln=True, align='C')
    pdf.ln(10)
    
    # Table Header
    pdf.set_font("Arial", style="B", size=12)
    pdf.cell(100, 10, txt="Name", border=1, align='C')
    pdf.cell(40, 10, txt="Points", border=1, align='C')
    pdf.ln()
    
    # Table Rows
    pdf.set_font("Arial", size=12)
    for name, points in employees_sorted:
        pdf.cell(100, 10, txt=name, border=1, align='L')
        pdf.cell(40, 10, txt=str(points), border=1, align='C')
        pdf.ln()
    
    # Page Number
    pdf.alias_nb_pages()
    pdf.set_y(-15)
    pdf.set_font("Arial", size=8)
    pdf.cell(0, 10, txt=f"Page {pdf.page_no()} of {{nb}}", align='C')

    pdf.output(output_path)

# Paths
input_pdf_path = r'C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\CallInsReport.pdf'
output_pdf_path = r'C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\SortedEmployeePoints.pdf'

# Execution
pdf_text = extract_data_from_pdf(input_pdf_path)
employee_data = parse_employee_data(pdf_text)
create_sorted_pdf(employee_data, output_pdf_path)

print(f"Sorted PDF created at {output_pdf_path}")
