import pdfplumber
import pandas as pd
import re

pdf_path = r"C:\Users\SamuelJoshuaRaj\Downloads\1254_CYGNUSA TECHNOLOGIES PVT LTD.pdf"

rows = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if not text:
            continue
        
        lines = text.split("\n")
        
        for line in lines:
            # Keep only lines that start with numbers (actual data rows)
            if re.match(r'^\d+', line):
                parts = line.split()
                
                # Filter rows with enough columns (your table has many numbers)
                if len(parts) >= 15:
                    rows.append(parts)

# Convert to DataFrame
df = pd.DataFrame(rows)


# Save to Excel
df.to_excel(r"C:\Users\SamuelJoshuaRaj\OneDrive - CYGNUSA Technologies\Desktop\Balaji invoices\invoice_clean.xlsx", index=False)

print("✅ Data extracted successfully to invoice_clean.xlsx")