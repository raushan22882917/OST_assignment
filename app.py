from flask import Flask, render_template, request, redirect
import os
import docx
from PyPDF2 import PdfReader
import openpyxl
import pandas as pd
import re


app = Flask(__name__)

def extract_text_from_docx(docx_path):
    doc = docx.Document(docx_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        text = ''
        for page_num in range(len(reader.pages)):
            text += reader.pages[page_num].extract_text()
        return text

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        folder_path = request.form['folder_path']
        file_data = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.docx'):
                    file_path = os.path.join(root, file)
                    text = extract_text_from_docx(file_path)
                    file_data.append((file, text))
                elif file.endswith('.pdf'):
                    file_path = os.path.join(root, file)
                    text = extract_text_from_pdf(file_path)
                    file_data.append((file, text))
        
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(['File Name', 'Text Content'])
        for file_name, text_content in file_data:
            worksheet.append([file_name, text_content])
        
        output_folder_path = os.path.join(folder_path, 'K:/INTERNDATA')
        os.makedirs(output_folder_path, exist_ok=True)
        excel_file_path = os.path.join(output_folder_path, 'file_data.xlsx')
        workbook.save(excel_file_path)
        
        return redirect('/')
    return render_template('index.html')


@app.route('/generate_cv_details', methods=['POST'])
def generate_cv_details():
    # Read the Excel file
    data = pd.read_excel("file_data.xlsx")

    # Define regex patterns
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_number_pattern = r'\b\d{10}\b'

    html_output = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CV Details</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            width: 1000px;
            margin: 0 auto;
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
            grid-gap: 20px;
        }
        .card {
            background-color: #fff;
            border-radius: 5px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            padding: 20px;
            height: 200px;
            overflow: auto;
        }
        .card h3 {
            margin-top: 0;
        }
        .card p {
            margin-bottom: 10px;
        }
    </style>
    </head>
    <body>
    <div class="container">
    """

    # Iterate over each row
    for index, row in data.iterrows():
        filename = row['File Name']
        text = row['Text Content']
        
        # Extract email
        email_match = re.search(email_pattern, text)
        email = email_match.group() if email_match else None
        
        # Extract phone number
        phone_number_match = re.search(phone_number_pattern, text)
        phone_number = phone_number_match.group() if phone_number_match else None
        
        # Construct HTML card
        html_card = f"""
        <div class="card">
            <h3>Filename: {filename}</h3>
            <p>Email: {email}</p>
            <p>Phone Number: {phone_number}</p>
            <hr>
            <p>{text}</p>
        </div>
        """
        
        # Append HTML card to the output
        html_output += html_card

    html_output += """
    </div>
    </body>
    </html>
    """

    return html_output


