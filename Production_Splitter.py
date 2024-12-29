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


