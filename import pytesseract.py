import pytesseract
from pdf2image import convert_from_path
import os
from tqdm import tqdm

pdf_path = r"C:\Users\SamuelJoshuaRaj\OneDrive - CYGNUSA Technologies\Desktop\elections\merged_output.pdf"
poppler_path = r"C:\Program Files (x86)\poppler-25.12.0\Library\bin"

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


images = convert_from_path(pdf_path, poppler_path=poppler_path)

text_output = []

for i, image in enumerate(tqdm(images)):
    text = pytesseract.image_to_string(image, lang='eng')
    text_output.append(text)

with open("output.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(text_output))

print("OCR Completed")