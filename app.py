import os
import smtplib
import time
import threading
import queue
import pandas as pd
import numpy as np
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
from flask import Flask, render_template, request, jsonify, send_file
from io import BytesIO
import base64
import json
import logging
from logging.handlers import RotatingFileHandler
import pickle
from datetime import datetime
import re
from bs4 import BeautifulSoup
from werkzeug.middleware.proxy_fix import ProxyFix

# Initialize Flask app
app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

# Configuration
class Config:
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    MAX_MESSAGE_SIZE = 25 * 1024 * 1024  # 25MB
    MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024  # 25MB
    DELAY_BETWEEN_EMAILS = 1  # seconds
    MAX_THREADS = 3  # Reduced to prevent timeouts
    DELAY_BETWEEN_BATCHES = 5  # seconds between thread batches
    PROGRESS_DIR = 'progress_data'
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max upload size
    TIMEOUT = 300  # 5 minutes for long operations

app.config.from_object(Config)

# Configure logging
def setup_logging():
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            RotatingFileHandler('logs/bulk_emailer.log', maxBytes=5*1024*1024, backupCount=3),
            logging.StreamHandler()
        ]
    )

setup_logging()
logger = logging.getLogger(__name__)

# Custom JSON encoder for numpy/pandas types
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

# Helper functions
def validate_email(email):
    """Validate email format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def validate_excel_file(file):
    """Validate Excel file content"""
    try:
        df = pd.read_excel(file)
        if 'Email' not in df.columns:
            return False, "'Email' column is required"
        if len(df) == 0:
            return False, "Excel file is empty"
        return True, ""
    except Exception as e:
        return False, f"Invalid Excel file: {str(e)}"

def validate_html_content(html):
    """Basic HTML validation"""
    try:
        BeautifulSoup(html, 'html.parser')
        return True
    except:
        return False

def calculate_message_size(msg):
    """Calculate the approximate size of the email message in bytes"""
    return len(msg.as_string().encode('utf-8'))

def save_progress(session_id, data):
    """Save sending progress to a file"""
    if not os.path.exists(app.config['PROGRESS_DIR']):
        os.makedirs(app.config['PROGRESS_DIR'])
    
    filename = os.path.join(app.config['PROGRESS_DIR'], f'{session_id}.pkl')
    with open(filename, 'wb') as f:
        pickle.dump(data, f)

def load_progress(session_id):
    """Load sending progress from file"""
    filename = os.path.join(app.config['PROGRESS_DIR'], f'{session_id}.pkl')
    if os.path.exists(filename):
        with open(filename, 'rb') as f:
            return pickle.load(f)
    return None

def clear_progress(session_id):
    """Clear progress data after completion"""
    filename = os.path.join(app.config['PROGRESS_DIR'], f'{session_id}.pkl')
    if os.path.exists(filename):
        os.remove(filename)

def validate_smtp_credentials(email, password):
    """Validate SMTP credentials by attempting to login"""
    try:
        with smtplib.SMTP(app.config['SMTP_SERVER'], app.config['SMTP_PORT'], timeout=10) as server:
            server.starttls()
            server.login(email, password)
            return True
    except Exception as e:
        logger.error(f"SMTP validation failed: {str(e)}")
        return False

def send_single_email(server, from_email, row, subject_template, html_template, attachments, report_data, error_log):
    """Send a single email with proper error handling"""
    try:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = row['Email']
        msg['Subject'] = subject_template

        # Personalize HTML content
        personalized_html = html_template
        for col, val in row.items():
            personalized_html = personalized_html.replace(f"{{{col}}}", str(val))

        msg.attach(MIMEText(personalized_html, 'html'))

        # Check message size before attachments
        if calculate_message_size(msg) > app.config['MAX_MESSAGE_SIZE']:
            raise ValueError("Message too large before attachments")

        # Handle PDF attachments from Excel
        if 'PDF_Path' in row and pd.notna(row['PDF_Path']):
            if os.path.exists(row['PDF_Path']):
                with open(row['PDF_Path'], 'rb') as f:
                    part = MIMEApplication(f.read(), _subtype='pdf')
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(row['PDF_Path']))
                    msg.attach(part)

                if calculate_message_size(msg) > app.config['MAX_MESSAGE_SIZE']:
                    raise ValueError("Message exceeded size limit after adding PDF attachment")

        # Handle other attachments
        for attachment in attachments:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment.filename}')
            msg.attach(part)
            attachment.seek(0)

            if calculate_message_size(msg) > app.config['MAX_MESSAGE_SIZE']:
                raise ValueError("Message exceeded size limit after adding attachment")

        # Final size check
        if calculate_message_size(msg) > app.config['MAX_MESSAGE_SIZE']:
            raise ValueError("Message too large to send")

        # Send the email
        server.sendmail(from_email, row['Email'], msg.as_string())
        time.sleep(app.config['DELAY_BETWEEN_EMAILS'])

        report_data.append({
            'Email': row['Email'],
            'Status': 'Success',
            'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        })

    except Exception as e:
        error_log.append({
            'email': row['Email'],
            'error': str(e),
            'help_link': 'https://support.google.com/mail/?p=MaxSizeError' if 'Message too large' in str(e) else None
        })
        report_data.append({
            'Email': row['Email'],
            'Status': 'Failed',
            'Error': str(e),
            'Timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        })

# Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/validate_smtp', methods=['POST'])
def validate_smtp():
    email = request.form.get('email')
    password = request.form.get('password')
    
    if not email or not password:
        return jsonify({'valid': False, 'message': 'Email and password are required'})
    
    if not validate_email(email):
        return jsonify({'valid': False, 'message': 'Invalid email format'})
    
    try:
        if validate_smtp_credentials(email, password):
            return jsonify({'valid': True, 'message': 'SMTP credentials are valid'})
        return jsonify({'valid': False, 'message': 'Invalid SMTP credentials'})
    except Exception as e:
        logger.error(f"Error validating SMTP: {str(e)}")
        return jsonify({'valid': False, 'message': f'Error validating credentials: {str(e)}'})

@app.route('/analyze_excel', methods=['POST'])
def analyze_excel():
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'No file provided'})

        excel_file = request.files['file']
        if not excel_file or excel_file.filename == '':
            return jsonify({'status': 'error', 'message': 'No file selected'})

        is_valid, message = validate_excel_file(excel_file)
        if not is_valid:
            return jsonify({'status': 'error', 'message': message})

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
        logger.error(f"Error analyzing Excel: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/send_test_email', methods=['POST'])
def send_test_email():
    try:
        data = request.json
        if not data:
            return jsonify({'status': 'error', 'message': 'No data provided'})
            
        email = data.get('email')
        smtp_email = data.get('smtp_email')
        smtp_password = data.get('smtp_password')
        subject = data.get('subject')
        template = data.get('template')
        
        # Check if all required fields are present
        if not all([email, smtp_email, smtp_password, subject, template]):
            return jsonify({'status': 'error', 'message': 'Missing required fields'})
            
        if not validate_email(email):
            return jsonify({'status': 'error', 'message': 'Invalid recipient email'})
            
        if not validate_smtp_credentials(smtp_email, smtp_password):
            return jsonify({'status': 'error', 'message': 'Invalid SMTP credentials'})
            
        server = smtplib.SMTP(app.config['SMTP_SERVER'], app.config['SMTP_PORT'], timeout=10)
        server.starttls()
        server.login(smtp_email, smtp_password)
        
        msg = MIMEMultipart()
        msg['From'] = smtp_email
        msg['To'] = email
        msg['Subject'] = "Test Email: " + subject
        
        if not validate_html_content(template):
            return jsonify({'status': 'error', 'message': 'Invalid HTML content'})
            
        msg.attach(MIMEText(template, 'html'))
        
        server.sendmail(smtp_email, email, msg.as_string())
        server.quit()
        
        return jsonify({'status': 'success', 'message': 'Test email sent successfully'})
        
    except Exception as e:
        logger.error(f"Error sending test email: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/send_emails', methods=['POST'])
def send_emails():
    try:
        session_id = str(datetime.now().timestamp())
        logger.info(f"Starting email sending session {session_id}")

        required_fields = ['smtp_email', 'smtp_password', 'subject', 'template']
        if not all(field in request.form for field in required_fields):
            return jsonify({'status': 'error', 'message': 'Missing required fields'})

        smtp_email = request.form['smtp_email']
        smtp_password = request.form['smtp_password']
        subject_template = request.form['subject']
        html_template = request.form['template']

        if not validate_html_content(html_template):
            return jsonify({'status': 'error', 'message': 'Invalid HTML template'})

        excel_file = request.files.get('excel_file')
        if not excel_file:
            return jsonify({'status': 'error', 'message': 'No Excel file provided'})

        # Check for existing progress
        progress_data = load_progress(session_id)
        if progress_data:
            df = progress_data['df']
            sent_count = progress_data['sent_count']
            report_data = progress_data['report_data']
            error_log = progress_data['error_log']
            logger.info(f"Resuming from progress: {sent_count} emails already sent")
        else:
            df = pd.read_excel(excel_file)
            df = df.dropna(subset=['Email'])
            df = df[df['Email'].str.contains('@', na=False)]
            df = df.fillna('')
            sent_count = 0
            report_data = []
            error_log = []

        attachments = [f for f in request.files.values() if f.filename and f != excel_file]

        # Validate attachment sizes
        for attachment in attachments:
            attachment.seek(0, os.SEEK_END)
            size = attachment.tell()
            attachment.seek(0)
            if size > app.config['MAX_ATTACHMENT_SIZE']:
                return jsonify({
                    'status': 'error',
                    'message': f'Attachment {attachment.filename} exceeds maximum size of 25MB',
                    'help_link': 'https://support.google.com/mail/?p=MaxSizeError'
                })

        # Connect to SMTP server
        server = smtplib.SMTP(app.config['SMTP_SERVER'], app.config['SMTP_PORT'], timeout=10)
        server.starttls()
        server.login(smtp_email, smtp_password)

        # Threaded email sending with reduced batch size
        email_queue = queue.Queue()
        for _, row in df.iterrows():
            email_queue.put(row)

        # Worker function with proper nonlocal declaration
        def email_worker():
            nonlocal sent_count, report_data, error_log
            while not email_queue.empty():
                try:
                    row = email_queue.get()
                    row_position = df.index.get_indexer([row.name])[0]
                    if row_position < sent_count:
                        email_queue.task_done()
                        continue

                    send_single_email(server, smtp_email, row, subject_template, 
                                    html_template, attachments, report_data, error_log)
                    
                    # Update sent count and save progress periodically
                    sent_count += 1
                    if sent_count % 5 == 0:  # Save more frequently
                        save_progress(session_id, {
                            'df': df,
                            'sent_count': sent_count,
                            'report_data': report_data,
                            'error_log': error_log
                        })
                        
                    email_queue.task_done()
                except Exception as e:
                    logger.error(f"Error sending email to {row['Email']}: {str(e)}")
                    email_queue.task_done()

        # Start threads in smaller batches
        total_emails = len(df)
        while sent_count < total_emails:
            threads = []
            for _ in range(min(app.config['MAX_THREADS'], total_emails - sent_count)):
                t = threading.Thread(target=email_worker)
                t.start()
                threads.append(t)

            # Wait for this batch to complete
            for t in threads:
                t.join(timeout=app.config['TIMEOUT'])
                
            # Delay between batches
            if sent_count < total_emails:
                time.sleep(app.config['DELAY_BETWEEN_BATCHES'])

        server.quit()
        clear_progress(session_id)

        # Generate report
        report_df = pd.DataFrame(report_data)
        report_buffer = BytesIO()
        with pd.ExcelWriter(report_buffer, engine='openpyxl') as writer:
            report_df.to_excel(writer, index=False, sheet_name='Email Report')
        report_buffer.seek(0)
        report_base64 = base64.b64encode(report_buffer.getvalue()).decode('utf-8')

        return jsonify({
            'status': 'success',
            'message': f'Successfully sent {sent_count} out of {total_emails} emails',
            'sentCount': sent_count,
            'totalCount': total_emails,
            'errors': error_log,
            'report': report_base64
        })

    except Exception as e:
        logger.error(f"Error in send_emails: {str(e)}", exc_info=True)
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