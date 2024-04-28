# app.py
import os
import shutil
import re
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file
import textract
from docx import Document
import PyPDF2

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.')[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
    return text

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

def extract_text_from_doc(doc_path):
    text = textract.process(doc_path, extension='doc').decode('utf-8')
    return text

def extract_contact_info(text):
    email_pattern = r'[\w\.-]+@[\w\.-]+'
    phone_pattern = r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})'
    
    email = re.search(email_pattern, text)
    phone = re.search(phone_pattern, text)
    
    email = email.group(0) if email else "Not found"
    phone = phone.group(0) if phone else "Not found"
    
    return email, phone

def process_cv_files(directory):
    data = []
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            text = extract_text_from_pdf(os.path.join(directory, filename))
        elif filename.endswith(".docx"):
            text = extract_text_from_docx(os.path.join(directory, filename))
        elif filename.endswith(".doc"):
            text = extract_text_from_doc(os.path.join(directory, filename))
        else:
            continue
        
        email, phone = extract_contact_info(text)
        data.append({'Filename': filename, 'Email': email, 'Phone': phone, 'Text': text})
    
    df = pd.DataFrame(data)
    return df

@app.route('/', methods=['GET', 'POST'])
def upload_cv():

    # deletes previous .xls file
    if os.path.exists(os.path.join(app.config['UPLOAD_FOLDER'], 'cv_data.xlsx')):
        os.remove(os.path.join(app.config['UPLOAD_FOLDER'], 'cv_data.xlsx'))

    # delete previously uploaded files
    files = os.listdir(app.config['UPLOAD_FOLDER'])

    for file_name in files:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
        if os.path.isfile(file_path):
            os.remove(file_path)
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path)
    
    # saves the uploaded files
    if request.method == 'POST':
        uploaded_files = request.files.getlist('file')
        for file in uploaded_files:
            if file.filename == '':
                continue
            if file:
                filename = file.filename
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    return render_template('index.html')

@app.route('/process', methods=['GET', 'POST'])
def process():
    cv_directory = app.config['UPLOAD_FOLDER']
    cv_data = process_cv_files(cv_directory)
    output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'cv_data.xlsx')
    cv_data.to_excel(output_file, index=False)
    return redirect(url_for('download'))

@app.route('/download')
def download():
    filename = os.path.join(app.config['UPLOAD_FOLDER'], 'cv_data.xlsx')
    return send_file(filename, as_attachment=True)
    
if __name__ == "__main__":
    app.run(port=10000)
