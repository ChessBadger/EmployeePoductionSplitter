import pdfplumber
import pandas as pd
import re
from datetime import datetime, timedelta

# File path
pdf_path = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\EmployeeProductionByStoreReport.pdf"
csv_output_path = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\EmployeeProduction.csv"

# Initialize an empty list to hold the extracted data
data = []

# Redefine parsing with cross-page employee tracking
last_employee = None  # Tracks the last employee name across pages

def parse_employee_data_with_carryover(page_text):
    global last_employee
    lines = page_text.split("\n")
    for line in lines:
        # Skip report date and page footer lines
        if re.match(r"\w+day, \w+ \d{1,2}, \d{4} Page \d+ of \d+", line):
            continue

        # Check for an employee name (no date or numeric values)
        if not any(char.isdigit() for char in line) and line.strip():
            last_employee = line.strip()  # Update the current employee name
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
                        "Employee": last_employee,
                        "Date": date,
                        "Store": store,
                        "Pieces/Hr": pieces_hr,
                        "$/Hr": dollars_hr,
                        "Skus/Hr": skus_hr,
                    })
                except ValueError:
                    continue

# Extract text from the PDF with cross-page tracking
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        parse_employee_data_with_carryover(text)

# Convert the data to a DataFrame
df = pd.DataFrame(data)

# Clean and process the DataFrame
df["Employee"] = df["Employee"].str.replace("Pieces/Hr \\$/Hr Skus/Hr", "", regex=True)
df["Employee"] = df["Employee"].fillna(method="ffill")  # Carry forward any missing employee names

# Convert the 'Date' column to datetime
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

# Calculate three-month threshold
three_months_ago = datetime.now() - timedelta(days=90)

# Identify employees with records within the last three months
active_employees = df[df["Date"] >= three_months_ago]["Employee"].unique()

# Filter the DataFrame to include only active employees
df = df[df["Employee"].isin(active_employees)]

# Export to CSV
df.to_csv(csv_output_path, index=False)
print(f"Data has been exported to {csv_output_path}")


import pandas as pd
from fpdf import FPDF
import os

# Paths
csv_path = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\EmployeeProduction.csv"
output_dir = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\EmployeeProductionReports"

# Create output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Load the CSV into a DataFrame
df = pd.read_csv(csv_path)

# Clean and convert columns to numeric
df["Pieces/Hr"] = pd.to_numeric(df["Pieces/Hr"].str.replace(",", "", regex=True), errors="coerce").fillna(0)
df["$/Hr"] = pd.to_numeric(df["$/Hr"].str.replace("[\$,]", "", regex=True), errors="coerce").fillna(0)
df["Skus/Hr"] = pd.to_numeric(df["Skus/Hr"].astype(str).str.replace(",", "", regex=True), errors="coerce").fillna(0)

# Group data by employee
grouped = df.groupby("Employee")

# Define a function to create PDF
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Employee Performance Report", align="C", ln=True)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")

    def truncate_text(self, text, max_length=30):
        """Truncates text to a maximum length, adding ellipsis if necessary."""
        if len(text) > max_length:
            return text[:max_length - 3] + "..."
        return text

    def add_table_row(self, col1, col2, col3, col4, col5):
        self.set_font("Arial", size=10)
        self.cell(30, 10, col1, border=1, align="C")
        self.cell(70, 10, col2, border=1, align="L")
        self.cell(30, 10, col3, border=1, align="C")
        self.cell(30, 10, col4, border=1, align="C")
        self.cell(30, 10, col5, border=1, align="C")
        self.ln()

for employee, group in grouped:
    pdf = PDF()
    pdf.add_page()

    # Title for the employee
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Employee: {employee}", ln=True)
    pdf.ln(5)

    # Add table header
    pdf.set_font("Arial", "B", 10)
    pdf.cell(30, 10, "Date", border=1, align="C")
    pdf.cell(70, 10, "Store", border=1, align="C")
    pdf.cell(30, 10, "Pieces/Hr", border=1, align="C")
    pdf.cell(30, 10, "$/Hr", border=1, align="C")
    pdf.cell(30, 10, "Skus/Hr", border=1, align="C")
    pdf.ln()

    # Add table rows
    pdf.set_font("Arial", size=10)
    for index, row in group.iterrows():
        store = pdf.truncate_text(row["Store"], max_length=30)  # Truncate the Store column
        pdf.add_table_row(
            row["Date"],
            store,
            f"{row['Pieces/Hr']:.2f}" if row["Pieces/Hr"] > 0 else "N/A",
            f"{row['$/Hr']:.2f}" if row["$/Hr"] > 0 else "N/A",
            f"{row['Skus/Hr']:.2f}" if row["Skus/Hr"] > 0 else "N/A",
        )

    # Filter out blank/0 values before calculating averages
    avg_pieces = group.loc[group["Pieces/Hr"] > 0, "Pieces/Hr"].mean()
    avg_dollars = group.loc[group["$/Hr"] > 0, "$/Hr"].mean()
    avg_skus = group.loc[group["Skus/Hr"] > 0, "Skus/Hr"].mean()

    # Add averages section
    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Overall Averages", ln=True)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 10, f"Pieces/Hr: {avg_pieces:.2f}" if pd.notna(avg_pieces) else "Pieces/Hr: No Data", ln=True)
    pdf.cell(0, 10, f"$/Hr: {avg_dollars:.2f}" if pd.notna(avg_dollars) else "$/Hr: No Data", ln=True)
    pdf.cell(0, 10, f"Skus/Hr: {avg_skus:.2f}" if pd.notna(avg_skus) else "Skus/Hr: No Data", ln=True)

    # Save PDF for this employee
    pdf_file_path = os.path.join(output_dir, f"{employee.replace(' ', '_')}.pdf")
    pdf.output(pdf_file_path)

print(f"PDFs created in {output_dir}")

import pandas as pd
from fpdf import FPDF
import os

# Paths
csv_path = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\EmployeeProduction.csv"
summary_pdf_path = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\SummaryReport.pdf"

# Load the CSV into a DataFrame
df = pd.read_csv(csv_path)

# Clean and convert columns to numeric
df["Pieces/Hr"] = pd.to_numeric(df["Pieces/Hr"].str.replace(",", "", regex=True), errors="coerce").fillna(0)
df["$/Hr"] = pd.to_numeric(df["$/Hr"].str.replace("[\$,]", "", regex=True), errors="coerce").fillna(0)
df["Skus/Hr"] = pd.to_numeric(df["Skus/Hr"].astype(str).str.replace(",", "", regex=True), errors="coerce").fillna(0)

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

    def add_summary_row(self, employee, pieces, dollars, skus, grand_avg_pieces, grand_avg_dollars, grand_avg_skus):
        self.set_font("Arial", size=10)
        
        # Employee name
        self.cell(70, 10, employee, border=1, align="L")

        # Pieces/Hr
        if pd.notna(pieces) and pieces < grand_avg_pieces:
            self.set_text_color(255, 0, 0)  # Red
            self.set_font("Arial", "BU", 10)  # Bold and Underline
        self.cell(40, 10, f"{pieces:.2f}" if pd.notna(pieces) else "N/A", border=1, align="C")
        self.set_text_color(0, 0, 0)  # Reset to black
        self.set_font("Arial", size=10)  # Reset to normal font

        # $/Hr
        if pd.notna(dollars) and dollars < grand_avg_dollars:
            self.set_text_color(255, 0, 0)  # Red
            self.set_font("Arial", "BU", 10)  # Bold and Underline
        self.cell(40, 10, f"{dollars:.2f}" if pd.notna(dollars) else "N/A", border=1, align="C")
        self.set_text_color(0, 0, 0)  # Reset to black
        self.set_font("Arial", size=10)  # Reset to normal font

        # Skus/Hr
        if pd.notna(skus) and skus < grand_avg_skus:
            self.set_text_color(255, 0, 0)  # Red
            self.set_font("Arial", "BU", 10)  # Bold and Underline
        self.cell(40, 10, f"{skus:.2f}" if pd.notna(skus) else "N/A", border=1, align="C")
        self.set_text_color(0, 0, 0)  # Reset to black
        self.set_font("Arial", size=10)  # Reset to normal font

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
    pdf.add_summary_row(
        employee=row["Employee"],
        pieces=row["Avg Pieces/Hr"],
        dollars=row["Avg $/Hr"],
        skus=row["Avg Skus/Hr"],
        grand_avg_pieces=grand_avg_pieces,
        grand_avg_dollars=grand_avg_dollars,
        grand_avg_skus=grand_avg_skus,
    )



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
