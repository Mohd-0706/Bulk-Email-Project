from flask import Flask, render_template, request, jsonify, send_file
import smtplib
import time
import pandas as pd
import numpy as np
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import json
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
import base64
import zipfile
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)

# Constants
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
MAX_MESSAGE_SIZE = 25 * 1024 * 1024  # Google's 25MB limit
MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024  # 25MB per attachment
DELAY_BETWEEN_EMAILS = 1  # seconds

# Helper function to calculate message size
def calculate_message_size(msg):
    """Calculate the approximate size of the email message in bytes"""
    return len(msg.as_string().encode('utf-8'))

# Compress large attachments to zip
def compress_attachment(file_path):
    """Compress attachment using ZIP format"""
    zip_path = file_path + '.zip'
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(file_path, os.path.basename(file_path))
    return zip_path

# Validate SMTP credentials
def validate_smtp_credentials(email, password):
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(email, password)
            return True
    except Exception as e:
        app.logger.error(f"SMTP validation failed: {str(e)}")
        return False

# Sending Email with Attachments
def send_email_with_attachment(smtp_email, smtp_password, row, subject_template, html_template, attachments, df):
    try:
        # Prepare Email
        msg = MIMEMultipart()
        msg['From'] = smtp_email
        msg['To'] = row['Email']
        msg['Subject'] = subject_template

        personalized_html = html_template
        for column in df.columns:
            personalized_html = personalized_html.replace(f"{{{column}}}", str(row[column]))

        msg.attach(MIMEText(personalized_html, 'html'))

        # Check message size before adding attachments
        if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
            return {'email': row['Email'], 'status': 'Failed', 'error': 'Message too large before attachments'}

        for attachment in attachments:
            # Compress if the attachment is large
            if attachment.size > MAX_ATTACHMENT_SIZE:
                file_path = attachment.filename
                compressed_file = compress_attachment(file_path)
                attachment = open(compressed_file, 'rb')

            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment.filename}')
            msg.attach(part)
            attachment.seek(0)

        # Final size check before sending
        if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
            return {'email': row['Email'], 'status': 'Failed', 'error': 'Message too large after attachments'}

        # Send Email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(smtp_email, smtp_password)
            server.sendmail(smtp_email, row['Email'], msg.as_string())

        return {'email': row['Email'], 'status': 'Success'}

    except Exception as e:
        return {'email': row['Email'], 'status': 'Failed', 'error': str(e)}

# Main route
@app.route('/')
def index():
    return render_template('index.html')

# Route to validate SMTP credentials
@app.route('/validate_smtp', methods=['POST'])
def validate_smtp():
    email = request.form.get('email')
    password = request.form.get('password')
    
    if not email or not password:
        return jsonify({'valid': False, 'message': 'Email and password are required'})
    
    try:
        if validate_smtp_credentials(email, password):
            return jsonify({'valid': True, 'message': 'SMTP credentials are valid'})
        else:
            return jsonify({'valid': False, 'message': 'Invalid SMTP credentials'})
    except Exception as e:
        return jsonify({'valid': False, 'message': f'Error validating credentials: {str(e)}'})

# Analyze Excel file and show preview
@app.route('/analyze_excel', methods=['POST'])
def analyze_excel():
    try:
        excel_file = request.files['file']
        if not excel_file:
            return jsonify({'status': 'error', 'message': 'No file provided'})

        df = pd.read_excel(excel_file)
        df = df.dropna(subset=['Email'])
        df = df[df['Email'].str.contains('@', na=False)]
        df = df.fillna('')

        preview_data = df.head(5).to_dict(orient='records')

        return json.dumps({
            'status': 'success',
            'columns': df.columns.tolist(),
            'rowCount': len(df),
            'preview': preview_data
        }, cls=NpEncoder), 200, {'Content-Type': 'application/json'}

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

# Route to send emails with attachments
@app.route('/send_emails', methods=['POST'])
def send_emails():
    try:
        required_fields = ['smtp_email', 'smtp_password', 'subject', 'template']
        if not all(field in request.form for field in required_fields):
            return jsonify({'status': 'error', 'message': 'Missing required fields'})

        smtp_email = request.form['smtp_email']
        smtp_password = request.form['smtp_password']
        subject_template = request.form['subject']
        html_template = request.form['template']

        excel_file = request.files.get('excel_file')
        if not excel_file:
            return jsonify({'status': 'error', 'message': 'No Excel file provided'})

        df = pd.read_excel(excel_file)
        df = df.dropna(subset=['Email'])
        df = df[df['Email'].str.contains('@', na=False)]
        df = df.fillna('')

        attachments = [f for f in request.files.values() if f.filename and f != excel_file]

        # Use ThreadPoolExecutor to send emails in parallel
        with ThreadPoolExecutor(max_workers=5) as executor:
            email_results = list(executor.map(
                send_email_with_attachment,
                [smtp_email] * len(df),
                [smtp_password] * len(df),
                df.to_dict(orient='records'),
                [subject_template] * len(df),
                [html_template] * len(df),
                [attachments] * len(df),
                [df] * len(df)
            ))

        sent_count = sum(1 for result in email_results if result['status'] == 'Success')
        error_log = [result for result in email_results if result['status'] == 'Failed']

        return jsonify({
            'status': 'success',
            'message': f'Successfully sent {sent_count} emails.',
            'sentCount': sent_count,
            'errorLog': error_log
        })

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

# Download sample Excel template
@app.route('/download_sample')
def download_sample():
    sample_data = [
        {'Email': 'example1@example.com', 'Name': 'John Doe', 'Company': 'ABC Corp', 'PDF_Path': ''},
        {'Email': 'example2@example.com', 'Name': 'Jane Smith', 'Company': 'XYZ Inc', 'PDF_Path': ''},
        {'Email': 'example3@example.com', 'Name': 'Bob Johnson', 'Company': '123 Industries', 'PDF_Path': ''}
    ]

    df = pd.DataFrame(sample_data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Recipients', index=False)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='BulkMail_Sample_Template.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True)
