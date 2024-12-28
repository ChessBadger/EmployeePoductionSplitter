import os
from fpdf import FPDF
import pdfplumber
import pandas as pd
import re

def extract_data_from_pdf(pdf_path):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue  # Skip pages without text
            lines = text.split("\n")
            current_name = None

            for line in lines:
                if "Pieces/Hr" in line and "Store" not in line:
                    current_name = line.split("Pieces/Hr")[0].strip()
                elif current_name and "Pieces/Hr" not in line:
                    parts = line.split()
                    if len(parts) > 4:  # At least Date, Store, ..., $/Hr, Pieces/Hr, SKUs/Hr
                        try:
                            date = parts[0]
                            pieces_index = len(parts) - 3
                            store = " ".join(parts[1:pieces_index])
                            pieces_hr = parts[pieces_index]
                            data.append({
                                "Name": current_name,
                                "Date": date,
                                "Store": store,
                                "Pieces/Hr": pieces_hr,
                            })
                        except IndexError:
                            print(f"Failed to parse line: {line}")
    return data

def clean_dataframe(df):
    valid_date_regex = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")
    cleaned_df = df[
        df["Date"].apply(lambda x: bool(valid_date_regex.match(str(x)))) &
        df["Pieces/Hr"].apply(lambda x: x.replace(',', '').isdigit())
    ]
    cleaned_df["Pieces/Hr"] = cleaned_df["Pieces/Hr"].str.replace(',', '').astype(int)
    return cleaned_df

def generate_employee_pdfs(df, output_dir):
    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for name, group in df.groupby("Name"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        
        # Title
        pdf.cell(200, 10, f"Employee Report: {name}", ln=True, align='C')
        
        # Table Header
        pdf.cell(200, 10, "Store - Pieces/Hr", ln=True, align='L')
        
        # Store Data
        total_pieces = 0
        count = 0
        for _, row in group.iterrows():
            pdf.cell(200, 10, f"{row['Store']} - {row['Pieces/Hr']} pieces/hr", ln=True, align='L')
            total_pieces += row['Pieces/Hr']
            count += 1
        
        # Average Pieces/Hr
        if count > 0:
            average_pieces = total_pieces / count
            pdf.cell(200, 10, f"\nOverall Average: {average_pieces:.2f} pieces/hr", ln=True, align='L')
        
        # Save PDF
        pdf_file = os.path.join(output_dir, f"{name.replace(' ', '_')}_report.pdf")
        pdf.output(pdf_file)
        print(f"Generated PDF for {name}: {pdf_file}")

def generate_ranked_pdf(df, output_file):
    # Calculate averages
    employee_averages = df.groupby("Name")["Pieces/Hr"].mean().sort_values(ascending=False)
    overall_average = df["Pieces/Hr"].mean()

    # Create PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Title
    pdf.cell(200, 10, "Employee Ranking by Average Pieces/Hr", ln=True, align='C')
    pdf.ln(10)

    # Rankings
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, "Rank - Name - Average Pieces/Hr", ln=True, align='L')
    rank = 1
    for name, avg in employee_averages.items():
        pdf.cell(200, 10, f"{rank}. {name} - {avg:.2f}", ln=True, align='L')
        rank += 1

    pdf.ln(10)
    # Overall average
    pdf.set_font("Arial", style="B", size=12)
    pdf.cell(200, 10, f"Overall Average: {overall_average:.2f} Pieces/Hr", ln=True, align='C')

    # Save PDF
    pdf.output(output_file)
    print(f"Generated ranking PDF: {output_file}")

if __name__ == "__main__":
    # File paths
    pdf_path = r"C:\Users\Laptop 122\Desktop\Store Prep\EmployeeProductionByStoreReport.pdf"
    output_dir = r"C:\Users\Laptop 122\Desktop\Store Prep\EmployeeReports"
    ranking_pdf_path = r"C:\Users\Laptop 122\Desktop\Store Prep\EmployeeRanking.pdf"

    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Step 1: Extract data
    extracted_data = extract_data_from_pdf(pdf_path)

    # Step 2: Clean data
    df = pd.DataFrame(extracted_data)
    cleaned_df = clean_dataframe(df)

    # Step 3: Generate PDFs for each employee
    generate_employee_pdfs(cleaned_df, output_dir)

    # Step 4: Generate a ranking PDF
    generate_ranked_pdf(cleaned_df, ranking_pdf_path)
