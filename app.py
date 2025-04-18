from flask import Flask, render_template, request, jsonify, send_file
import smtplib
import time
import pandas as pd
import numpy as np
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import json
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import io

app = Flask(__name__)

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Custom JSON encoder to handle numpy and pandas types
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
        return super(NpEncoder, self).default(obj)

@app.route('/')
def index():
    return render_template('index.html')

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

        # Process additional attachments
        attachments = [f for f in request.files.values() if f.filename and f.name.startswith('attachment_')]

        for _, row in df.iterrows():
            try:
                msg = MIMEMultipart()
                msg['From'] = smtp_email
                msg['To'] = row['Email']
                msg['Subject'] = subject_template

                # Personalize content
                personalized_html = html_template
                for column in df.columns:
                    personalized_html = personalized_html.replace(f"{{{column}}}", str(row[column]))

                msg.attach(MIMEText(personalized_html, 'html'))

                # Attach PDF if specified
                if 'PDF_Path' in df.columns and pd.notna(row['PDF_Path']):
                    pdf_path = row['PDF_Path']
                    if os.path.exists(pdf_path):
                        with open(pdf_path, 'rb') as f:
                            part = MIMEApplication(f.read(), _subtype='pdf')
                            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
                            msg.attach(part)

                # Attach additional files
                for attachment in attachments:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={attachment.filename}')
                    msg.attach(part)
                    attachment.seek(0)

                server.sendmail(smtp_email, row['Email'], msg.as_string())
                sent_count += 1
                time.sleep(1)  # optional rate limiting

            except Exception as email_err:
                continue  # Skip to next row on failure

        server.quit()

        return jsonify({
            'status': 'success',
            'message': f'Successfully sent {sent_count} out of {total_count} emails',
            'sentCount': sent_count,
            'totalCount': total_count
        })

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/download_sample')
def download_sample():
    sample_data = [
        {'Email': 'example1@example.com', 'Name': 'John Doe', 'Company': 'ABC Corp', 'PDF_Path': ''},
        {'Email': 'example2@example.com', 'Name': 'Jane Smith', 'Company': 'XYZ Inc', 'PDF_Path': ''},
        {'Email': 'example3@example.com', 'Name': 'Bob Johnson', 'Company': '123 Industries', 'PDF_Path': ''}
    ]

    df = pd.DataFrame(sample_data)

    output = io.BytesIO()
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