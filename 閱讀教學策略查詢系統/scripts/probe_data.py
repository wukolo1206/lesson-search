
import os
import pdfplumber
import pandas as pd
import json

def list_files(directory):
    return [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.pdf')]

def scan_folders():
    folders = [d for d in os.listdir('.') if os.path.isdir(d) and '學習扶助' in d]
    for folder in folders:
        print(f"Folder: {folder}")
        files = list_files(folder)
        for f in files:
            print(f"  File: {f}")

def extract_sample_text(pdf_path, page_num=10):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) > page_num:
                text = pdf.pages[page_num].extract_text()
                print(f"--- Text from {pdf_path} (Page {page_num}) ---\n{text}\n--- End ---")
                return text
            else:
                print(f"{pdf_path} has only {len(pdf.pages)} pages.")
    except Exception as e:
        print(f"Error opening {pdf_path}: {e}")

if __name__ == "__main__":
    # Detect the correct folder paths
    folders = os.listdir('.')
    supp_folder = next((f for f in folders if '補充文本' in f), None)
    curriculum_folder = next((f for f in folders if '國語教材' in f), None)
    
    print(f"Detected Supplementary Folder: {supp_folder}")
    print(f"Detected Curriculum Folder: {curriculum_folder}")
    
    if supp_folder:
        supp_files = list_files(supp_folder)
        if supp_files:
            # Try to extract from the "補救教學讀本1-6年級.pdf" or similar
            sample_file = next((f for f in supp_files if '1-6' in f), supp_files[0])
            extract_sample_text(sample_file, 20)
    
    if curriculum_folder:
        curr_files = list_files(curriculum_folder)
        if curr_files:
            sample_file = next((f for f in curr_files if '3-3-2' in f), curr_files[0])
            extract_sample_text(sample_file, 5)
