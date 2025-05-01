from flask import Flask, render_template, request, jsonify, send_file
import smtplib
import time
import pandas as pd
import numpy as np
import os
import json
import zipfile
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)

# Constants
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
MAX_MESSAGE_SIZE = 25 * 1024 * 1024  # 25MB total size limit
MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024  # 25MB per file
DELAY_BETWEEN_EMAILS = 1  # seconds (optional throttle)


# JSON encoder for NumPy types
class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.integer,)):
            return int(obj)
        elif isinstance(obj, (np.floating,)):
            return float(obj)
        elif isinstance(obj, (np.ndarray,)):
            return obj.tolist()
        return super(NpEncoder, self).default(obj)


# Calculate message size
def calculate_message_size(msg):
    return len(msg.as_string().encode('utf-8'))


# Compress large attachment to ZIP
def compress_attachment(file_path):
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


# Send personalized email
def send_email_with_attachment(smtp_email, smtp_password, row, subject_template, html_template, attachments, df):
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_email
        msg['To'] = row['Email']
        msg['Subject'] = subject_template

        # Replace placeholders with actual values
        personalized_html = html_template
        for column in df.columns:
            personalized_html = personalized_html.replace(f"{{{column}}}", str(row[column]))
        msg.attach(MIMEText(personalized_html, 'html'))

        # Initial size check
        if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
            return {'email': row['Email'], 'status': 'Failed', 'error': 'Message too large before attachments'}

        for attachment in attachments:
            file_data = attachment.read()
            file_size = len(file_data)

            # Compress if large
            if file_size > MAX_ATTACHMENT_SIZE:
                temp_path = f"/tmp/{attachment.filename}"
                with open(temp_path, 'wb') as f:
                    f.write(file_data)
                compressed_path = compress_attachment(temp_path)
                with open(compressed_path, 'rb') as f:
                    file_data = f.read()
                filename = os.path.basename(compressed_path)
            else:
                filename = attachment.filename

            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file_data)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            msg.attach(part)

        # Final size check
        if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
            return {'email': row['Email'], 'status': 'Failed', 'error': 'Message too large after attachments'}

        # Send
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(smtp_email, smtp_password)
            server.sendmail(smtp_email, row['Email'], msg.as_string())

        time.sleep(DELAY_BETWEEN_EMAILS)
        return {'email': row['Email'], 'status': 'Success'}

    except Exception as e:
        return {'email': row['Email'], 'status': 'Failed', 'error': str(e)}


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/validate_smtp', methods=['POST'])
def validate_smtp():
    email = request.form.get('email')
    password = request.form.get('password')

    if not email or not password:
        return jsonify({'valid': False, 'message': 'Email and password are required'})

    if validate_smtp_credentials(email, password):
        return jsonify({'valid': True, 'message': 'SMTP credentials are valid'})
    else:
        return jsonify({'valid': False, 'message': 'Invalid SMTP credentials'})


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


@app.route('/send_emails', methods=['POST'])
def send_emails():
    try:
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

        # Send in parallel
        with ThreadPoolExecutor(max_workers=5) as executor:
            results = list(executor.map(
                send_email_with_attachment,
                [smtp_email] * len(df),
                [smtp_password] * len(df),
                df.to_dict(orient='records'),
                [subject_template] * len(df),
                [html_template] * len(df),
                [attachments] * len(df),
                [df] * len(df)
            ))

        sent_count = sum(1 for r in results if r['status'] == 'Success')
        error_log = [r for r in results if r['status'] == 'Failed']

        return jsonify({
            'status': 'success',
            'message': f'Successfully sent {sent_count} emails.',
            'sentCount': sent_count,
            'errorLog': error_log
        })

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/download_sample')
def download_sample():
    sample_data = [
        {'Email': 'example1@example.com', 'Name': 'John Doe', 'Company': 'ABC Corp'},
        {'Email': 'example2@example.com', 'Name': 'Jane Smith', 'Company': 'XYZ Inc'},
        {'Email': 'example3@example.com', 'Name': 'Bob Johnson', 'Company': '123 Industries'}
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
