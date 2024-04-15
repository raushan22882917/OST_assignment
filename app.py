import os
import win32com.client
from docx import Document
from PyPDF2 import PdfReader
import openpyxl
import pandas as pd
import re
from flask import Flask, render_template, request, redirect
import pythoncom

app = Flask(__name__)

processed_files = set()  # To keep track of processed file paths
processed_emails = set()  # To keep track of processed email addresses

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
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

def convert_doc_to_docx(doc_file):
    pythoncom.CoInitialize()  # Initialize COM
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_file)
    docx_file = doc_file.replace(".doc", ".docx")
    doc.SaveAs(docx_file, FileFormat=16)  # FileFormat 16 represents .docx
    doc.Close()
    word.Quit()
    return docx_file

def clean_excel_file(excel_file_path):
    if os.path.exists(excel_file_path):
        data = pd.read_excel(excel_file_path)
        # Drop duplicate rows based on email address
        data.drop_duplicates(subset=['Email'], inplace=True)
        # Remove unwanted files (e.g., temporary files)
        data = data[~data['File Name'].str.startswith('~$')]  # Remove files starting with ~$
        data = data[~data['File Name'].str.startswith('.')]   # Remove hidden files
        data = data[~data['File Name'].str.endswith('.tmp')]  # Remove temporary files
        # Save cleaned data back to the Excel file
        data.to_excel(excel_file_path, index=False)

def extract_email_and_phone(text):
    # Define regex patterns
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_number_pattern = r'\b\d{10}\b'

    # Extract potential email addresses and phone numbers
    potential_emails = re.findall(email_pattern, text)
    potential_phone_numbers = re.findall(phone_number_pattern, text)

    return potential_emails, potential_phone_numbers

def clean_email(email):
    # Remove spaces from email and ensure it's in a valid format
    email = email.replace(" ", "")
    return email

def clean_phone_number(phone_number):
    # Remove spaces from phone number and ensure it's in a valid format
    phone_number = phone_number.replace(" ", "")
    return phone_number

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        folder_path = request.form['folder_path']
        excel_file_path = os.path.join(folder_path, 'K:/INTERNDATA', 'file_data.xlsx')
        # Clear old data from the Excel file
        if os.path.exists(excel_file_path):
            os.remove(excel_file_path)
        file_data = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                if file_path not in processed_files:  # Check if file already processed
                    if file.endswith('.docx'):
                        text = extract_text_from_docx(file_path)
                        email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
                        phone_number_match = re.search(r'\b\d{10}\b', text)
                        if email_match:
                            email = email_match.group()
                        else:
                            potential_emails, _ = extract_email_and_phone(text)
                            email = clean_email(potential_emails[0]) if potential_emails else None
                        if email not in processed_emails:
                            if phone_number_match:
                                phone_number = phone_number_match.group()
                            else:
                                _, potential_phone_numbers = extract_email_and_phone(text)
                                phone_number = clean_phone_number(potential_phone_numbers[0]) if potential_phone_numbers else None
                            file_data.append((file, text, email, phone_number))
                            processed_emails.add(email)
                    elif file.endswith('.doc'):
                        docx_file_path = convert_doc_to_docx(file_path)
                        text = extract_text_from_docx(docx_file_path)
                        email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
                        phone_number_match = re.search(r'\b\d{10}\b', text)
                        if email_match:
                            email = email_match.group()
                        else:
                            potential_emails, _ = extract_email_and_phone(text)
                            email = clean_email(potential_emails[0]) if potential_emails else None
                        if email not in processed_emails:
                            if phone_number_match:
                                phone_number = phone_number_match.group()
                            else:
                                _, potential_phone_numbers = extract_email_and_phone(text)
                                phone_number = clean_phone_number(potential_phone_numbers[0]) if potential_phone_numbers else None
                            file_data.append((os.path.basename(docx_file_path), text, email, phone_number))  # Use docx file name
                            processed_emails.add(email)
                    elif file.endswith('.pdf'):
                        text = extract_text_from_pdf(file_path)
                        email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
                        phone_number_match = re.search(r'\b\d{10}\b', text)
                        if email_match:
                            email = email_match.group()
                        else:
                            potential_emails, _ = extract_email_and_phone(text)
                            email = clean_email(potential_emails[0]) if potential_emails else None
                        if email not in processed_emails:
                            if phone_number_match:
                                phone_number = phone_number_match.group()
                            else:
                                _, potential_phone_numbers = extract_email_and_phone(text)
                                phone_number = clean_phone_number(potential_phone_numbers[0]) if potential_phone_numbers else None
                            file_data.append((file, text, email, phone_number))
                            processed_emails.add(email)
                    processed_files.add(file_path)  # Add the processed file path
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(['File Name', 'Text Content', 'Email', 'Mobile Number'])
        for file_name, text_content, email, mobile_number in file_data:
            worksheet.append([file_name, text_content, email, mobile_number])
        output_folder_path = os.path.join(folder_path, 'K:/INTERNDATA')
        os.makedirs(output_folder_path, exist_ok=True)
        clean_excel_file(excel_file_path)
        workbook.save(excel_file_path)
        return redirect('/')
    return render_template('index.html')

@app.route('/generate_cv_details', methods=['POST'])
def generate_cv_details():
    excel_file_path = os.path.join('K:/INTERNDATA', 'file_data.xlsx')
    clean_excel_file(excel_file_path)
    # Read the cleaned Excel file
    data = pd.read_excel(excel_file_path)

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
            background-color: #f4f4f4;
            background-image: linear-gradient(to right, #77A1D3 0%, #79CBCA  51%, #77A1D3  100%);
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
            cursor: pointer; /* Add cursor pointer */
        }
        .card h3 {
            margin-top: 0;
            cursor: pointer; /* Add cursor pointer */
        }
        .card p {
            margin-bottom: 10px;
        }
        .large-card { /* Add style for large card */
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 800px;
            height: 600px;
            background-color: #fff;
            border-radius: 5px;
            padding: 20px;
            overflow-y: auto;
            z-index: 1000;
        }
        .close-btn {
            position: absolute;
            top: 10px;
            right: 10px;
            cursor: pointer;
            color: #555;
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
        email = row['Email']
        mobile_number = row['Mobile Number']
        
        # Format text with headings and bullet points
        formatted_text = "<h4>Text Content:</h4>"
        formatted_text += "<ul>"
        for line in text.split('\n'):
            if line.strip():  # Only include non-empty lines
                formatted_text += f"<li>{line}</li>"
        formatted_text += "</ul>"

        # Construct HTML card
        html_card = f"""
        <div class="card" onclick="showLargeCard(this)">
            <h3>Filename: {filename}</h3>
            <p>Email: {email}</p>
            <p>Mobile Number: {mobile_number}</p>
            <hr>
            {formatted_text}
        </div>
        """
        
        # Append HTML card to the output
        html_output += html_card

    html_output += """
    </div>
    <div id="largeCardContainer" style="display: none;">
        <div id="largeCard" class="large-card"></div>
    </div>
    <script>
        function showLargeCard(card) {
            var largeCardContainer = document.getElementById('largeCardContainer');
            var largeCard = document.getElementById('largeCard');
            largeCard.innerHTML = card.innerHTML;
            largeCardContainer.style.display = 'block';
        }
        function closeLargeCard() {
            var largeCardContainer = document.getElementById('largeCardContainer');
            largeCardContainer.style.display = 'none';
        }
        document.getElementById('largeCard').addEventListener('click', function(event) {
            event.stopPropagation();
        });
        document.addEventListener('click', function() {
            closeLargeCard();
        });
    </script>
    </body>
    </html>
    """

    return html_output

if __name__ == '__main__':
    app.run(debug=True)
