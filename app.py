from flask import Flask, render_template, request, jsonify, send_file
import smtplib
import pandas as pd
import numpy as np
import os
import time
import json
import base64
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)

# Email sending configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
MAX_MESSAGE_SIZE = 25 * 1024 * 1024  # 25 MB
MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024
DELAY_BETWEEN_EMAILS = 1  # seconds


# JSON Encoder for pandas/numpy
class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer): return int(obj)
        if isinstance(obj, np.floating): return float(obj)
        if isinstance(obj, np.ndarray): return obj.tolist()
        if pd.isna(obj): return None
        return super().default(obj)


def calculate_message_size(msg):
    return len(msg.as_string().encode('utf-8'))


def validate_smtp_credentials(email, password):
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(email, password)
        return True
    except Exception as e:
        app.logger.error(f"SMTP validation failed: {str(e)}")
        return False


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
        excel_file.seek(0)
        df = pd.read_excel(excel_file)
        df = df.dropna(subset=['Email'])
        df = df[df['Email'].str.contains('@', na=False)].fillna('')

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
        # Validate fields
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

        excel_file.seek(0)
        df = pd.read_excel(excel_file)
        df = df.dropna(subset=['Email'])
        df = df[df['Email'].str.contains('@', na=False)].fillna('')

        attachments = [f for f in request.files.values() if f.filename and f != excel_file]

        # Check attachment sizes
        for attachment in attachments:
            attachment.seek(0, os.SEEK_END)
            if attachment.tell() > MAX_ATTACHMENT_SIZE:
                return jsonify({'status': 'error', 'message': f'Attachment {attachment.filename} exceeds 25MB'})
            attachment.seek(0)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(smtp_email, smtp_password)

        sent_count = 0
        error_log = []
        report_data = []

        for _, row in df.iterrows():
            try:
                msg = MIMEMultipart()
                msg['From'] = smtp_email
                msg['To'] = row['Email']
                msg['Subject'] = subject_template

                body = html_template
                for col in df.columns:
                    body = body.replace(f"{{{col}}}", str(row[col]))
                msg.attach(MIMEText(body, 'html'))

                # Add PDF from path if specified
                if 'PDF_Path' in df.columns and pd.notna(row['PDF_Path']):
                    pdf_path = row['PDF_Path']
                    if os.path.exists(pdf_path):
                        with open(pdf_path, 'rb') as f:
                            part = MIMEApplication(f.read(), _subtype='pdf')
                            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
                            msg.attach(part)

                # Add uploaded attachments
                for attachment in attachments:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={attachment.filename}')
                    msg.attach(part)
                    attachment.seek(0)

                if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
                    error_log.append({'email': row['Email'], 'error': 'Message too large (>25MB)'})
                    report_data.append({'Email': row['Email'], 'Status': 'Failed', 'Error': 'Size > 25MB'})
                    continue

                server.sendmail(smtp_email, row['Email'], msg.as_string())
                sent_count += 1
                report_data.append({'Email': row['Email'], 'Status': 'Success'})
                time.sleep(DELAY_BETWEEN_EMAILS)

            except Exception as e:
                error_log.append({'email': row['Email'], 'error': str(e)})
                report_data.append({'Email': row['Email'], 'Status': 'Failed', 'Error': str(e)})

        server.quit()

        # Return report
        report_df = pd.DataFrame(report_data)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            report_df.to_excel(writer, index=False)
        buffer.seek(0)

        report_b64 = base64.b64encode(buffer.read()).decode('utf-8')

        return jsonify({
            'status': 'success',
            'sentCount': sent_count,
            'totalCount': len(df),
            'errors': error_log,
            'report': report_b64,
            'message': f"{sent_count}/{len(df)} emails sent successfully."
        })

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/download_sample')
def download_sample():
    data = [
        {'Email': 'test@example.com', 'Name': 'John', 'Company': 'XYZ Inc', 'PDF_Path': ''},
        {'Email': 'jane@example.com', 'Name': 'Jane', 'Company': 'ABC Ltd', 'PDF_Path': ''}
    ]
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Recipients')
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Sample_Template.xlsx'
    )


if __name__ == '__main__':
    app.run(debug=True)
