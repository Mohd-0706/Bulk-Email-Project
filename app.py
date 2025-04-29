from flask import Flask, render_template, request, jsonify, send_file
import smtplib
import time
import pandas as pd
import numpy as np
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import json
import logging
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
import base64

app = Flask(__name__)

# Constants
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
MAX_MESSAGE_SIZE = 25 * 1024 * 1024  # Google's 25MB limit
MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024  # 25MB per attachment
DELAY_BETWEEN_EMAILS = 1  # seconds

class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        if pd.isna(obj):
            return None
        return super().default(obj)

def validate_smtp_credentials(email, password):
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(email, password)
            return True
    except Exception as e:
        app.logger.error(f"SMTP validation failed: {str(e)}")
        return False

def calculate_message_size(msg):
    """Calculate the approximate size of the email message in bytes"""
    return len(msg.as_string().encode('utf-8'))

@app.route('/')
def index():
    return render_template('index.html')

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

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(smtp_email, smtp_password)

        sent_count = 0
        total_count = len(df)
        error_log = []
        report_data = []

        attachments = [f for f in request.files.values() if f.filename and f != excel_file]

        # Validate attachment sizes before sending
        for attachment in attachments:
            attachment.seek(0, os.SEEK_END)
            size = attachment.tell()
            attachment.seek(0)
            if size > MAX_ATTACHMENT_SIZE:
                server.quit()
                return jsonify({
                    'status': 'error',
                    'message': f'Attachment {attachment.filename} exceeds maximum size of 25MB',
                    'help_link': 'https://support.google.com/mail/?p=MaxSizeError'
                })

        for _, row in df.iterrows():
            try:
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
                    error_log.append({
                        'email': row['Email'],
                        'error': 'Message too large before attachments',
                        'size': calculate_message_size(msg)
                    })
                    report_data.append({
                        'Email': row['Email'],
                        'Status': 'Failed',
                        'Error': 'Message too large before attachments',
                        'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    continue

                if 'PDF_Path' in df.columns and pd.notna(row['PDF_Path']):
                    pdf_path = row['PDF_Path']
                    if os.path.exists(pdf_path):
                        with open(pdf_path, 'rb') as f:
                            part = MIMEApplication(f.read(), _subtype='pdf')
                            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
                            msg.attach(part)

                            # Check message size after each attachment
                            if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
                                error_log.append({
                                    'email': row['Email'],
                                    'error': 'Message exceeded size limit after adding PDF attachment',
                                    'size': calculate_message_size(msg)
                                })
                                report_data.append({
                                    'Email': row['Email'],
                                    'Status': 'Failed',
                                    'Error': 'Message exceeded size limit',
                                    'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                                })
                                continue

                for attachment in attachments:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={attachment.filename}')
                    msg.attach(part)
                    attachment.seek(0)

                    # Check message size after each attachment
                    if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
                        error_log.append({
                            'email': row['Email'],
                            'error': 'Message exceeded size limit after adding attachment',
                            'size': calculate_message_size(msg),
                            'help_link': 'https://support.google.com/mail/?p=MaxSizeError'
                        })
                        report_data.append({
                            'Email': row['Email'],
                            'Status': 'Failed',
                            'Error': 'Message exceeded 25MB size limit',
                            'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                        })
                        continue

                # Final size check before sending
                if calculate_message_size(msg) > MAX_MESSAGE_SIZE:
                    error_log.append({
                        'email': row['Email'],
                        'error': 'Message too large to send',
                        'size': calculate_message_size(msg),
                        'help_link': 'https://support.google.com/mail/?p=MaxSizeError'
                    })
                    report_data.append({
                        'Email': row['Email'],
                        'Status': 'Failed',
                        'Error': 'Message exceeded 25MB size limit',
                        'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    continue

                server.sendmail(smtp_email, row['Email'], msg.as_string())
                sent_count += 1
                report_data.append({
                    'Email': row['Email'],
                    'Status': 'Success',
                    'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                })
                time.sleep(DELAY_BETWEEN_EMAILS)

            except smtplib.SMTPDataError as e:
                if 'Message too large' in str(e):
                    error_log.append({
                        'email': row['Email'],
                        'error': 'Message exceeded Google size limit (25MB)',
                        'help_link': 'https://support.google.com/mail/?p=MaxSizeError'
                    })
                    report_data.append({
                        'Email': row['Email'],
                        'Status': 'Failed',
                        'Error': 'Message exceeded 25MB size limit',
                        'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                else:
                    error_log.append({
                        'email': row['Email'],
                        'error': str(e)
                    })
                    report_data.append({
                        'Email': row['Email'],
                        'Status': 'Failed',
                        'Error': str(e),
                        'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
            except Exception as e:
                error_log.append({
                    'email': row['Email'],
                    'error': str(e)
                })
                report_data.append({
                    'Email': row['Email'],
                    'Status': 'Failed',
                    'Error': str(e),
                    'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                })

        server.quit()

        # Generate report
        report_df = pd.DataFrame(report_data)
        report_buffer = BytesIO()
        with pd.ExcelWriter(report_buffer, engine='openpyxl') as writer:
            report_df.to_excel(writer, index=False, sheet_name='Email Report')
        report_buffer.seek(0)
        report_base64 = base64.b64encode(report_buffer.getvalue()).decode('utf-8')

        return jsonify({
            'status': 'success',
            'message': f'Successfully sent {sent_count} out of {total_count} emails',
            'sentCount': sent_count,
            'totalCount': total_count,
            'errors': error_log,
            'report': report_base64
        })

    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'help_link': 'https://support.google.com/mail/?p=MaxSizeError' if 'Message too large' in str(e) else None
        })

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