import pdfplumber
import pandas as pd
import re

# File path
pdf_path = r"C:\Users\Laptop 122\Desktop\Store Prep\EmployeeProductionByStoreReport.pdf"
csv_output_path = r"C:\Users\Laptop 122\Desktop\Store Prep\EmployeeProduction.csv"

# Initialize an empty list to hold the extracted data
data = []

# Define a function to parse the text and extract relevant fields
def parse_employee_data(page_text):
    lines = page_text.split("\n")
    current_employee = None  # Tracks the current employee's name

    for line in lines:
        # Skip report date and page footer lines
        if re.match(r"\w+day, \w+ \d{1,2}, \d{4} Page \d+ of \d+", line):
            continue

        # Check for an employee name (no date or numeric values)
        if not any(char.isdigit() for char in line) and line.strip():
            current_employee = line.strip()
        else:
            # Split the line into components
            parts = line.split()
            if len(parts) >= 2:
                try:
                    date = parts[0]  # First column is the date
                    store = []
                    pieces_hr = dollars_hr = skus_hr = ""

                    # Start parsing each part
                    for part in parts[1:]:
                        if part.startswith("$") and not dollars_hr:
                            # Found $/Hr
                            dollars_hr = part
                        elif part.replace(",", "").isdigit() and not pieces_hr:
                            # Found Pieces/Hr
                            pieces_hr = part
                        elif part.replace(",", "").isdigit() and not skus_hr:
                            # Found Skus/Hr
                            skus_hr = part
                        else:
                            # Add to store name
                            store.append(part)

                    # Combine store back into a single string
                    store = " ".join(store).strip()

                    # Append the row to the data
                    data.append({
                        "Employee": current_employee,
                        "Date": date,
                        "Store": store,
                        "Pieces/Hr": pieces_hr,
                        "$/Hr": dollars_hr,
                        "Skus/Hr": skus_hr,
                    })
                except ValueError:
                    continue

# Extract text from the PDF
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        parse_employee_data(text)

# Convert the data to a DataFrame
df = pd.DataFrame(data)

# Remove "Pieces/Hr $/Hr Skus/Hr" from every cell in the Employee column
df["Employee"] = df["Employee"].str.replace("Pieces/Hr \\$/Hr Skus/Hr", "", regex=True)

# Fill empty cells in the Employee column with the most recent non-empty value
df["Employee"] = df["Employee"].fillna(method="ffill")

# Export to CSV
df.to_csv(csv_output_path, index=False)
print(f"Data has been exported to {csv_output_path}")


import pandas as pd
from fpdf import FPDF
import os

# Paths
csv_path = r"C:\Users\Laptop 122\Desktop\Store Prep\EmployeeProduction.csv"
summary_pdf_path = r"C:\Users\Laptop 122\Desktop\Store Prep\SummaryReport.pdf"

# Load the CSV into a DataFrame
df = pd.read_csv(csv_path)

# Clean and convert columns to numeric
df["Pieces/Hr"] = pd.to_numeric(df["Pieces/Hr"].str.replace(",", "", regex=True), errors="coerce").fillna(0)
df["$/Hr"] = pd.to_numeric(df["$/Hr"].str.replace("[\$,]", "", regex=True), errors="coerce").fillna(0)
df["Skus/Hr"] = pd.to_numeric(df["Skus/Hr"].str.replace(",", "", regex=True), errors="coerce").fillna(0)

# Calculate cumulative averages for each employee
grouped = df.groupby("Employee")
employee_averages = []

for employee, group in grouped:
    avg_pieces = group.loc[group["Pieces/Hr"] > 0, "Pieces/Hr"].mean()
    avg_dollars = group.loc[group["$/Hr"] > 0, "$/Hr"].mean()
    avg_skus = group.loc[group["Skus/Hr"] > 0, "Skus/Hr"].mean()
    if pd.notna(avg_pieces) or pd.notna(avg_dollars) or pd.notna(avg_skus):
        employee_averages.append({
            "Employee": employee,
            "Avg Pieces/Hr": avg_pieces,
            "Avg $/Hr": avg_dollars,
            "Avg Skus/Hr": avg_skus,
        })

# Convert to DataFrame
averages_df = pd.DataFrame(employee_averages)

# Calculate grand averages
grand_avg_pieces = averages_df["Avg Pieces/Hr"].mean()
grand_avg_dollars = averages_df["Avg $/Hr"].mean()
grand_avg_skus = averages_df["Avg Skus/Hr"].mean()

# Sort employees by cumulative averages
averages_df["Cumulative Avg"] = averages_df[["Avg Pieces/Hr", "Avg $/Hr", "Avg Skus/Hr"]].mean(axis=1, skipna=True)
averages_df = averages_df.sort_values(by="Cumulative Avg", ascending=False)

# Create Summary PDF
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Employee Performance Summary", align="C", ln=True)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")

    def add_summary_row(self, employee, pieces, dollars, skus):
        self.set_font("Arial", size=10)
        self.cell(70, 10, employee, border=1, align="L")
        self.cell(40, 10, f"{pieces:.2f}" if pd.notna(pieces) else "N/A", border=1, align="C")
        self.cell(40, 10, f"{dollars:.2f}" if pd.notna(dollars) else "N/A", border=1, align="C")
        self.cell(40, 10, f"{skus:.2f}" if pd.notna(skus) else "N/A", border=1, align="C")
        self.ln()

pdf = PDF()
pdf.add_page()

# Add table header
pdf.set_font("Arial", "B", 10)
pdf.cell(70, 10, "Employee", border=1, align="C")
pdf.cell(40, 10, "Avg Pieces/Hr", border=1, align="C")
pdf.cell(40, 10, "Avg $/Hr", border=1, align="C")
pdf.cell(40, 10, "Avg Skus/Hr", border=1, align="C")
pdf.ln()

# Add employee rows
for _, row in averages_df.iterrows():
    pdf.add_summary_row(row["Employee"], row["Avg Pieces/Hr"], row["Avg $/Hr"], row["Avg Skus/Hr"])

# Add grand averages at the bottom
pdf.ln(10)
pdf.set_font("Arial", "B", 12)
pdf.cell(0, 10, "Grand Averages", ln=True)
pdf.set_font("Arial", size=10)
pdf.cell(0, 10, f"Avg Pieces/Hr: {grand_avg_pieces:.2f}" if pd.notna(grand_avg_pieces) else "Avg Pieces/Hr: No Data", ln=True)
pdf.cell(0, 10, f"Avg $/Hr: {grand_avg_dollars:.2f}" if pd.notna(grand_avg_dollars) else "Avg $/Hr: No Data", ln=True)
pdf.cell(0, 10, f"Avg Skus/Hr: {grand_avg_skus:.2f}" if pd.notna(grand_avg_skus) else "Avg Skus/Hr: No Data", ln=True)

# Save Summary PDF
pdf.output(summary_pdf_path)
print(f"Summary PDF created at {summary_pdf_path}")
