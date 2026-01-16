import os
import re
import pandas as pd
from pypdf import PdfReader


# Folder where your PDFs are stored
PDF_FOLDER = "C:\\Users\\Aayush\\OneDrive\\Desktop\\Invoice folder"

OUTPUT_FILE = "invoices_data.xlsx"

# Regex patterns for extraction
patterns = {
    "Invoice no": r"Invoice No\s+([A-Z0-9]+)",
    "Invoice date": r"Invoice Date\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})",
    "Passenger name": r"\n([A-Z\s]+)\nADULT",  # captures passenger name before "ADULT"
    "Travel date": r"(\d{2}-[A-Za-z]{3}-\d{4})\s*$",   # matches last date in the row (end of line)
    "Price": r"Total Amount\s+([\d,.]+)",
    "Price + Taxes": r"Total Due \(INR\)\s+([\d,.]+)"
}

def extract_invoice_data(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"

    data = {}
    data["File name"] = os.path.basename(pdf_path)

    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            data[key] = match.group(1).strip()
        else:
            data[key] = None

    # Extract Details: Airline + Ticket + Sector + Flight
    details_match = re.search(r"(Indigo Airlines.*?)(\d{2}-[A-Za-z]{3}-\d{4})", text, re.DOTALL)
    if details_match:
        data["Details"] = details_match.group(1).replace("\n", " ").strip()
    else:
        data["Details"] = None

    return data

# Process all PDFs in the folder
all_data = []
for file in os.listdir(PDF_FOLDER):
    if file.endswith(".pdf"):
        pdf_path = os.path.join(PDF_FOLDER, file)
        invoice_data = extract_invoice_data(pdf_path)
        all_data.append(invoice_data)

# Save to Excel
df = pd.DataFrame(all_data, columns=[
    "File name", "Invoice no", "Invoice date", 
    "Passenger name", "Details", "Travel date", 
    "Price", "Price + Taxes"
])
df.to_excel(OUTPUT_FILE, index=False)

print(f"Data extraction completed! Saved to {OUTPUT_FILE}")
