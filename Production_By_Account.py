import pandas as pd
from fpdf import FPDF
import os

# Helper function to truncate text so it doesn't overflow in the cell
def truncate_text(text, max_length=40):
    """
    Truncates a string to max_length characters, adding '...' if it's too long.
    """
    text = str(text)
    if len(text) <= max_length:
        return text
    else:
        return text[:max_length-3] + "..."

class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Production Report", ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

def main():
    # Mapping for account groups.
    # For example, "Kroger" expands to "pick", "mariano", and "metro"
    account_groups = {
        "kroger": ["pick", "mariano", "metro"],
        "f": ["pig", "festival", "sendik"]
        # Add more groups as needed
    }
    
    # Path to the CSV file
    file_path = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\EmployeeProduction.csv"
    
    # Read the CSV file into a DataFrame
    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"Error reading the CSV file: {e}")
        return

    # Clean and convert performance metric columns to numeric values
    df["Pieces/Hr"] = pd.to_numeric(df["Pieces/Hr"].str.replace(",", "", regex=True), errors="coerce")
    df["$/Hr"]      = pd.to_numeric(df["$/Hr"].str.replace("[\$,]", "", regex=True), errors="coerce")
    df["Skus/Hr"]   = pd.to_numeric(df["Skus/Hr"].str.replace(",", "", regex=True), errors="coerce")

    # Prompt user for multiple employee names and store substrings
    name_search_input = input("Enter employee names (comma-separated, leave blank for all): ").strip()
    store_search_input = input("Enter store substrings (comma-separated, leave blank for all): ").strip()

    # Convert inputs into lists
    name_search_list = [name.strip() for name in name_search_input.split(",") if name.strip()]
    store_search_list = [store.strip() for store in store_search_input.split(",") if store.strip()]
    
    # Expand store search list if any term matches a group name in our mapping
    expanded_store_search_list = []
    for term in store_search_list:
        key = term.lower()  # case-insensitive matching
        if key in account_groups:
            expanded_store_search_list.extend(account_groups[key])
        else:
            expanded_store_search_list.append(term)
    
    # Determine filtering logic
    if name_search_list:
        # If employee names are provided, filter by employee and optionally by store if provided
        filtered = df[
            df["Employee"].str.contains('|'.join(name_search_list), case=False, na=False) &
            (df["Store"].str.contains('|'.join(expanded_store_search_list), case=False, na=False) if expanded_store_search_list else True)
        ]
        group_by = "Store"
    else:
        # If no specific employee is provided, filter by store substrings only
        filtered = df[
            df["Store"].str.contains('|'.join(expanded_store_search_list), case=False, na=False) if expanded_store_search_list else True
        ]
        group_by = "Employee"

    if filtered.empty:
        print("No records found for the given search criteria.")
        return

    # Calculate overall averages for the filtered dataset
    avg_pieces_overall = filtered.loc[filtered["Pieces/Hr"] > 0, "Pieces/Hr"].mean()
    avg_dollars_overall = filtered.loc[filtered["$/Hr"] > 0, "$/Hr"].mean()
    avg_skus_overall = filtered.loc[filtered["Skus/Hr"] > 0, "Skus/Hr"].mean()

    # Prepare PDF
    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Arial", "", 12)
    
    # Display the search criteria at the top of the report
    pdf.cell(0, 10, f"Search Criteria: Employees: {', '.join(name_search_list) if name_search_list else 'All'}, Stores: {', '.join(expanded_store_search_list) if expanded_store_search_list else 'All'}", ln=True)
    pdf.ln(5)
    
    # Add the overall account averages section
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Account Averages:", ln=True)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Pieces/Hr: {avg_pieces_overall:.2f}" if pd.notna(avg_pieces_overall) else "Pieces/Hr: No Data", ln=True)
    pdf.cell(0, 10, f"$/Hr: {avg_dollars_overall:.2f}" if pd.notna(avg_dollars_overall) else "$/Hr: No Data", ln=True)
    pdf.cell(0, 10, f"Skus/Hr: {avg_skus_overall:.2f}" if pd.notna(avg_skus_overall) else "Skus/Hr: No Data", ln=True)
    pdf.ln(10)

    # Calculate individual averages by grouping by 'Employee'
    employee_averages = filtered.groupby("Employee").agg({
        "Pieces/Hr": "mean",
        "$/Hr": "mean",
        "Skus/Hr": "mean"
    }).reset_index()
    
    # Prompt the user for sorting preference using numeric choices
    sort_choice = input("Enter sort order for individual averages:\n1 - Alphabetical (by Employee)\n2 - Pieces/Hr (highest first)\n3 - $/Hr (highest first)\nYour choice: ").strip()
    if sort_choice == "1":
        employee_averages = employee_averages.sort_values(by="Employee", ascending=True)
    elif sort_choice == "2":
        employee_averages = employee_averages.sort_values(by="Pieces/Hr", ascending=False)
    elif sort_choice == "3":
        employee_averages = employee_averages.sort_values(by="$/Hr", ascending=False)
    else:
        print("Invalid choice, defaulting to alphabetical sorting.")
        employee_averages = employee_averages.sort_values(by="Employee", ascending=True)
    
    # Add a section for each individual's averages
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Individual Averages:", ln=True)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(50, 10, "Employee", border=1, align="C")
    pdf.cell(30, 10, "Pieces/Hr", border=1, align="C")
    pdf.cell(30, 10, "$/Hr", border=1, align="C")
    pdf.cell(30, 10, "Skus/Hr", border=1, align="C")
    pdf.ln()
    
    pdf.set_font("Arial", "", 10)
    for _, row in employee_averages.iterrows():
        pdf.cell(50, 10, truncate_text(row["Employee"], 25), border=1, align="C")
        pdf.cell(30, 10, f"{row['Pieces/Hr']:.2f}" if pd.notna(row['Pieces/Hr']) else "N/A", border=1, align="C")
        pdf.cell(30, 10, f"{row['$/Hr']:.2f}" if pd.notna(row['$/Hr']) else "N/A", border=1, align="C")
        pdf.cell(30, 10, f"{row['Skus/Hr']:.2f}" if pd.notna(row['Skus/Hr']) else "N/A", border=1, align="C")
        pdf.ln()
    pdf.ln(10)

    # Group and display detailed production data
    grouped = filtered.groupby(group_by)
    for key, group in grouped:
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, f"{group_by}: {key}", ln=True)

        # Table header
        pdf.set_font("Arial", "B", 10)
        pdf.cell(30, 10, "Date", border=1, align="C")
        pdf.cell(50, 10, "Employee" if group_by == "Store" else "Store", border=1, align="C")
        pdf.cell(30, 10, "Pieces/Hr", border=1, align="C")
        pdf.cell(30, 10, "$/Hr", border=1, align="C")
        pdf.cell(30, 10, "Skus/Hr", border=1, align="C")
        pdf.ln()

        # Table rows
        pdf.set_font("Arial", "", 10)
        for _, row in group.iterrows():
            pdf.cell(30, 10, str(row["Date"]), border=1, align="C")
            if group_by == "Store":
                # When grouped by Store, display Employee names (truncate at 25 chars)
                cell_text = truncate_text(row["Employee"], 25)
            else:
                # When grouped by Employee, display Store names (truncate at 15 chars to prevent overflow)
                cell_text = truncate_text(row["Store"], 20)
            pdf.cell(50, 10, cell_text, border=1, align="C")
            pdf.cell(30, 10, f"{row['Pieces/Hr']:.2f}" if pd.notna(row["Pieces/Hr"]) else "N/A", border=1, align="C")
            pdf.cell(30, 10, f"{row['$/Hr']:.2f}" if pd.notna(row["$/Hr"]) else "N/A", border=1, align="C")
            pdf.cell(30, 10, f"{row['Skus/Hr']:.2f}" if pd.notna(row["Skus/Hr"]) else "N/A", border=1, align="C")
            pdf.ln()


    # Generate output file name
    output_dir = r"C:\Users\Laptop 122\Desktop\Store Prep\06 Employee Reports\SearchedProductionReports"
    os.makedirs(output_dir, exist_ok=True)
    sanitized_name = "_".join(name_search_list).replace(" ", "_") if name_search_list else "All"
    sanitized_store = "_".join(expanded_store_search_list).replace(" ", "_") if expanded_store_search_list else "All"
    output_file = os.path.join(output_dir, f"{sanitized_name}_{sanitized_store}.pdf")

    # Save the PDF report
    pdf.output(output_file)
    print(f"PDF report created: {output_file}")

if __name__ == "__main__":
    main()
