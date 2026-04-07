import os
from concurrent.futures import ProcessPoolExecutor
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

# --- CONFIGURATION ---
PDF_PATH 
OUTPUT_DIR = r'C:\Users\SamuelJoshuaRaj\OneDrive - CYGNUSA Technologies\Desktop\PRASANTH'
# If Tesseract is not in your PATH, uncomment and point to the .exe
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def process_page(page_data):
    """Function to process a single page: OCR -> Text"""
    page_number, page_image = page_data
    try:
        # Perform OCR
        text = pytesseract.image_to_string(page_image)
        
        # Save individual page text to avoid losing progress
        output_file = os.path.join(OUTPUT_DIR, f"page_{page_number}.txt")
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(text)
            
        return f"Page {page_number} complete."
    except Exception as e:
        return f"Error on page {page_number}: {e}"

def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    print("Converting PDF to images... (This may take a while for 8000 pages)")
    
    # We use a generator to avoid loading 8000 images into RAM at once
    # 'thread_count' here is for the conversion process itself
    pages = convert_from_path(PDF_PATH, dpi=300, thread_count=4)

    print(f"Starting OCR with multiprocessing...")
    # Use ProcessPoolExecutor to use all available CPU cores
    with ProcessPoolExecutor() as executor:
        # Enumerate pages to keep track of page numbers
        results = list(executor.map(process_page, enumerate(pages, start=1)))

    print("\nProcessing finished. Check the 'extracted_pages' folder.")

if __name__ == "__main__":
    main()