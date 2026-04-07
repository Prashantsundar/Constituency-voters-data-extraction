import pdfplumber
import pandas as pd
import re

pdf_path = r"C:\Users\SamuelJoshuaRaj\Downloads\2026-EROLLGEN-S22-17-SIR-FinalRoll-Revision1-ENG-9-WI.pdf"

names = []
ages = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()

        if text:
            # extract name
            name_matches = re.findall(r"Name\s*[:\-]?\s*([^\n]+)", text)

            # extract age
            age_matches = re.findall(r"Age\s*[:\-]?\s*(\d+)", text)

            names.extend(name_matches)
            ages.extend(age_matches)

data = list(zip(names, ages))

df = pd.DataFrame(data, columns=["Name", "Age"])

print(df)

df.to_csv("output.csv", index=False)

print("Data extracted successfully!")