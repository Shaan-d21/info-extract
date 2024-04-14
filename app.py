from flask import Flask, render_template, request, send_file
from PyPDF2 import PdfReader
import re
import docx
import docx2txt
import os
from openpyxl import Workbook
import tempfile

app = Flask(__name__)

def extract_emails(text):
    # Extract emails from a given text.
    email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    emails = re.findall(email_pattern, text)
    return emails

def extract_phone_numbers(text):
    # Extract phone numbers from a given text.
    phone_pattern = r"\b\d{2,3}[-.\s]?\d{5,10}\b"
    phone_regex = re.compile(phone_pattern)
    matches = phone_regex.findall(text)
    return matches

def extract_text_from_pdf(file_path):
    # Extract text from a PDF file.
    with open(file_path, "rb") as f:
        reader = PdfReader(f)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_text_from_docx(file_path):
    # Extract text from a DOCX file.
    text = docx2txt.process(file_path)
    return text

def extract_text_from_doc(file_path):
    # Extract text from a DOC file.
    doc = docx.Document(file_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text
    return text

def extract_information(file_path):
    # Extract emails and phone numbers from a file.
    if file_path.endswith(".pdf"):
        text = extract_text_from_pdf(file_path)
    elif file_path.endswith(".docx"):
        text = extract_text_from_docx(file_path)
    elif file_path.endswith(".doc"):
        try:
            text = extract_text_from_doc(file_path)
        except docx.opc.exceptions.PackageNotFoundError:
            print(f"Error: Package not found for file '{file_path}'. Skipping...")
            text = ""
    else:
        raise ValueError("Unsupported file format")
    
    emails = extract_emails(text)
    phone_numbers = extract_phone_numbers(text)
    return text, emails, phone_numbers

def create_excel_file(files):
    # "Create and write to an Excel file.
    wb = Workbook()
    ws = wb.active
    ws.append(["File Name", "Text", "Emails", "Phone Numbers"])

    for file in files:
        file_name = file.filename
        file_path = os.path.join(tempfile.mkdtemp(), file_name)
        file.save(file_path)
        text, emails, phone_numbers = extract_information(file_path)
        ws.append([file_name, text, ", ".join(emails), ", ".join(phone_numbers)])

    excel_file_path = os.path.join(tempfile.mkdtemp(), "extracted_information.xlsx")
    wb.save(excel_file_path)
    return excel_file_path

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "files[]" not in request.files:
            return "No files part"
        files = request.files.getlist("files[]")
        if not files:
            return "No files selected"
        excel_file_path = create_excel_file(files)
        return send_file(excel_file_path, as_attachment=True)
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
