import streamlit as st
import fitz  # PyMuPDF
import re
import shutil
import os
import tempfile
import glob
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
from zipfile import ZipFile

st.title("PDF Decryption & Rename Tool")

# Temp directories
temp_dir = tempfile.mkdtemp()
pdf_folder_path = os.path.join(temp_dir, 'uploaded_pdfs')
output_folder_path = os.path.join(temp_dir, 'decrypted_pdfs')
renamed_folder_path = os.path.join(temp_dir, 'renamed_pdfs')

os.makedirs(pdf_folder_path, exist_ok=True)
os.makedirs(output_folder_path, exist_ok=True)
os.makedirs(renamed_folder_path, exist_ok=True)

logs = ""

def log(msg):
    global logs
    logs += msg + "\n"

# Step 1: Upload Excel with passwords
password_file = st.file_uploader("Upload Excel file with passwords (XLSX)", type=['xlsx'])

# Step 2: Upload multiple PDFs
uploaded_pdfs = st.file_uploader("Upload encrypted PDF files", type=['pdf'], accept_multiple_files=True)

if password_file and uploaded_pdfs:

    # Save uploaded PDFs to temp folder
    for pdf_file in uploaded_pdfs:
        file_path = os.path.join(pdf_folder_path, pdf_file.name)
        with open(file_path, 'wb') as f:
            f.write(pdf_file.read())
    log(f"Saved {len(uploaded_pdfs)} PDF files.")

    # Load passwords from Excel
    try:
        workbook = openpyxl.load_workbook(password_file)
        worksheet = workbook.active
        passwords = [str(cell.value) for cell in worksheet['A'] if cell.value]
        log(f"Loaded {len(passwords)} passwords from Excel.")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # Decrypt PDFs
    pdf_files = glob.glob(os.path.join(pdf_folder_path, '*.pdf'))
    for pdf_file in pdf_files:
        try:
            pdf_reader = PdfReader(pdf_file)
            if pdf_reader.is_encrypted:
                decrypted = False
                for pwd in passwords:
                    if pdf_reader.decrypt(pwd) > 0:
                        pdf_writer = PdfWriter()
                        for page in pdf_reader.pages:
                            pdf_writer.add_page(page)
                        file_name = os.path.splitext(os.path.basename(pdf_file))[0]
                        output_file_path = os.path.join(output_folder_path, file_name + '.pdf')
                        with open(output_file_path, 'wb') as output_file:
                            pdf_writer.write(output_file)
                        log(f"Decrypted and saved: {file_name}.pdf")
                        decrypted = True
                        break
                if not decrypted:
                    log(f"Failed to decrypt: {os.path.basename(pdf_file)}")
            else:
                log(f"PDF not encrypted: {os.path.basename(pdf_file)}")
        except Exception as e:
            log(f"Error decrypting {os.path.basename(pdf_file)}: {e}")

    # Function to extract numbers from 15th line
    def extract_numbers_from_15th_line(pdf_path):
        numbers_15th_line = []
        try:
            pdf_document = fitz.open(pdf_path)
            for page_num in range(pdf_document.page_count):
                page = pdf_document[page_num]
                lines = page.get_text().split('\n')
                if len(lines) > 14:
                    line_15 = lines[14].strip()
                    numbers = re.findall(r'\d+', line_15)
                    numbers_15th_line.extend(numbers)
            pdf_document.close()
        except Exception as e:
            log(f"Error extracting numbers from {os.path.basename(pdf_path)}: {e}")
        return numbers_15th_line

    # Rename and save PDFs
    decrypted_pdfs = glob.glob(os.path.join(output_folder_path, '*.pdf'))
    for dec_pdf in decrypted_pdfs:
        try:
            nums = extract_numbers_from_15th_line(dec_pdf)
            if nums:
                new_name = f"{nums[0]}.pdf"
                new_path = os.path.join(renamed_folder_path, new_name)
                shutil.copy2(dec_pdf, new_path)
                log(f"Renamed and saved: {new_name}")
            else:
                log(f"No numbers found in 15th line for {os.path.basename(dec_pdf)} - skipping rename.")
        except Exception as e:
            log(f"Error renaming {os.path.basename(dec_pdf)}: {e}")

    # Display all logs once
    st.text_area("Logs", value=logs, height=300, key="log_area")

    # Zip renamed files for download
    zip_path = os.path.join(temp_dir, "renamed_pdfs.zip")
    with ZipFile(zip_path, 'w') as zipf:
        for file in os.listdir(renamed_folder_path):
            zipf.write(os.path.join(renamed_folder_path, file), arcname=file)

    with open(zip_path, "rb") as fp:
        st.download_button(
            label="Download renamed PDFs ZIP",
            data=fp,
            file_name="renamed_pdfs.zip",
            mime="application/zip"
        )

else:
    st.info("Please upload both the Excel password file and at least one encrypted PDF file to start.")
