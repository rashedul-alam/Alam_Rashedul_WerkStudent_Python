import pdfplumber
import re
import pandas as pd
import os
from datetime import datetime

# Get the current directory
current_dir = os.getcwd()

# Keywords to search for
search_keywords = ["Gross Amount incl. VAT", "Total"]

# Function to convert date to mm.dd.yyyy format

def format_date(date_str):
    try:
        # Clean up the string
        date_str = date_str.replace("Invoice date:", "").strip()
        
        # German month name mapping
        german_months = {
            "Januar": "01", "Februar": "02", "März": "03", "April": "04", "Mai": "05", "Juni": "06",
            "Juli": "07", "August": "08", "September": "09", "Oktober": "10", "November": "11", "Dezember": "12"
        }

        # Try parsing German-style date
        if re.search(r"\d+\.\s*\w+\s*\d{4}", date_str):
            parts = date_str.split()
            day = parts[0].strip(".")  # Remove trailing period
            month = german_months.get(parts[1], None)  # Look up the month in the mapping
            year = parts[2]
            if month:
                formatted_date = f"{month}.{day.zfill(2)}.{year}"
                return formatted_date
            else:
                raise ValueError(f"Unknown German month in date: {date_str}")

        # Try parsing English-style date
        elif re.search(r"\w+\s+\d{1,2},\s+\d{4}", date_str):
            date_obj = datetime.strptime(date_str, "%b %d, %Y")
            return date_obj.strftime("%m.%d.%Y")

        # Return original string if no match is found
        return date_str
    except Exception as e:
        print(f"Error formatting date '{date_str}': {e}")
        return date_str

# Function to extract date and value from a PDF
def extract_data(pdf_path, keywords):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract the text as a structured layout of words and positions
            table_data = page.extract_table()
            
            # Check if the PDF page contains a table
            if table_data:
                # Search for the "Date" column
                for row in table_data:
                    for i, cell in enumerate(row):
                        if cell and "Date" in cell:
                            # Get the cell directly below or beside the "Date" column
                            date_value = row[i + 1] if i + 1 < len(row) else None
                            if date_value:
                                formatted_date = format_date(date_value.strip())
                                return formatted_date

            # Fallback to generic text extraction for keywords and other fields
            text = page.extract_text()
            if text:
                # Extract date as a fallback
                date_match = re.search(r"(\d+\.\s*\w+\s*\d{4})|Invoice date:\s*(\w+\s+\d{1,2},\s+\d{4})", text, re.IGNORECASE)
                date = format_date(date_match.group(0)) if date_match else None

                # Search for keywords
                for keyword in keywords:
                    keyword_pos = text.find(keyword)
                    if keyword_pos != -1:
                        # Extract value after the keyword
                        value_match = re.search(r"\d[\d,.]*", text[keyword_pos:])
                        if value_match:
                            value = float(value_match.group(0).replace(",", "."))
                            return date, keyword, value

    return None, None, None
# Extract data from all PDFs in the directory
data = []
for file_name in os.listdir(current_dir):
    if file_name.endswith(".pdf"):
        pdf_path = os.path.join(current_dir, file_name)
        date, keyword, value = extract_data(pdf_path, search_keywords)
        data.append({"File Name": file_name, "Date": date, "Keyword": keyword, "Value": value})

# Create a DataFrame
df = pd.DataFrame(data)

# Create a pivot table
pivot_table = pd.pivot_table(
    df,
    values="Value",
    index="Date",
    columns="File Name",
    aggfunc="sum",
    fill_value=0,
    margins=True,
    margins_name="Total"
)

# Save to Excel
output_excel = os.path.join(current_dir, "Invoice_Summery.xlsx")
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Sheet 1")
    pivot_table.to_excel(writer, sheet_name="Sheet 2")
# Save to CSV
output_csv = os.path.join(current_dir, "Invoice_Summery.csv")
df.to_csv(output_csv, sep=";", index=False)
