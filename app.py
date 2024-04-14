import os
import re
import pandas as pd
import PyPDF2
from docx import Document
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    cv_files = request.files.getlist('cv_files')
    if not cv_files:
        return 'Please select at least one file.', 400

    all_emails = []
    all_phones = []
    all_overall_text = []

    for cv_file in cv_files:
        if cv_file.filename.endswith('.pdf'):
            text = extract_text_from_pdf(cv_file)
        elif cv_file.filename.endswith('.docx'):
            text = extract_text_from_docx(cv_file)
        else:
            return f'Unsupported file format: {cv_file.filename}', 400

        emails, phones, overall_text = extract_info(text)
        all_emails.extend(emails)
        all_phones.extend(phones)
        all_overall_text.append(overall_text)

    data = {'Email': all_emails, 'Phone': all_phones, 'Overall Text': all_overall_text}
    file_name = "cv_info.xlsx"
    save_to_excel(data, file_name)

    # Construct the full file path
    file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), file_name)

    return send_file(file_path, as_attachment=True)

def extract_text_from_pdf(pdf_file):
    text = ""
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    for page_num in range(len(pdf_reader.pages)):
        text += pdf_reader.pages[page_num].extract_text().strip() + "\n"
    return text.strip()

def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    full_text = ""
    for para in doc.paragraphs:
        full_text += para.text.strip() + "\n"
    return full_text.strip()

def extract_info(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})'
    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)
    overall_text = text
    return emails, phones, overall_text

def save_to_excel(data, file_name):
    # Ensure all lists have the same length
    max_length = max(len(data['Email']), len(data['Phone']), len(data['Overall Text']))
    for key in data:
        if len(data[key]) < max_length:
            data[key].extend([''] * (max_length - len(data[key])))

    # Save the file to the same directory as the Flask app
    file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), file_name)
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)
    print(f"Information extracted and saved to {file_name}")

if __name__ == '__main__':
    app.run(debug=True)