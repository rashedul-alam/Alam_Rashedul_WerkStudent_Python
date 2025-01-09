
# WerkStudent_Python

## Overview

This script processes PDF invoices in a folder, extracts important information (like invoice dates and amounts), organizes the data, and outputs the results in Excel and CSV formats.

---

## **Features**

- Extracts **invoice dates** and **amounts** from PDF files.
- Handles **German** and **English** date formats.
- Processes tables or plain text in PDFs.
- Summarizes data into a pivot table grouped by date.
- Outputs:
  - A detailed data table.
  - A pivot table summary.
  - Saves results as **Excel** and **CSV** files.

---

## **Setup Instructions**

### **1. Prerequisites**

Ensure you have the following Python libraries installed:

- `pdfplumber`
- `pandas`
- `openpyxl`

### **2. Place Files**

1. Save the script in a folder.
2. Place your PDF files in the same folder.

---

## **How It Works**

1. **Define Folder and Keywords**
   - The script looks for PDF files in the same folder as the script.
   - Keywords like `Gross Amount incl. VAT` and `Total` are used to find relevant amounts.

2. **Extract Information**
   - The script uses `pdfplumber` to read each PDF.
   - It extracts:
     - **Dates**: Found in tables or plain text.
     - **Amounts**: Associated with the keywords.

3. **Format Dates**
   - The script uses `format_date(date_str)` to standardize dates into `mm.dd.yyyy` format.
   - Supports:
     - **German** date formats (e.g., `12. März 2023`).
     - **English** date formats (e.g., `March 12, 2023`).

4. **Organize Data**
   - Extracted data is stored in a Pandas DataFrame.
   - A pivot table summarizes the total amounts grouped by date.

5. **Save Outputs**
   - **Excel File**: Includes detailed data and summary.
   - **CSV File**: Contains detailed data.

## **Problem Reporting **
  -Due to GitHub's size restrictions, I could not upload the executable file directly to the repository.
