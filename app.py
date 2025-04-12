from flask import Flask, render_template, request, jsonify
import smtplib
import time
import pandas as pd
import numpy as np
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import json
from email.mime.application import MIMEApplication

app = Flask(__name__)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

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
            return jsonify({
                'status': 'error',
                'message': 'No file provided'
            })
            
        df = pd.read_excel(excel_file)
        
        # Clean data - remove empty rows and invalid emails
        df = df.dropna(subset=['Email'])
        df = df[df['Email'].str.contains('@', na=False)]
        df = df.fillna('')
        
        preview_data = []
        for _, row in df.head(5).iterrows():
            row_dict = {col: row[col] for col in df.columns}
            preview_data.append(row_dict)
        
        return json.dumps({
            'status': 'success',
            'columns': df.columns.tolist(),
            'rowCount': len(df),
            'preview': preview_data
        }, cls=NpEncoder), 200, {'Content-Type': 'application/json'}
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        })

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
        
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(smtp_email, smtp_password)
        
        sent_count = 0
        total_count = len(df)
        
        for index, row in df.iterrows():
            try:
                msg = MIMEMultipart()
                msg["From"] = smtp_email
                msg["To"] = str(row["Email"])
                msg["Subject"] = subject_template
                
                # Personalize HTML content
                personalized_html = html_template
                for column in df.columns:
                    placeholder = f"{{{column}}}"
                    personalized_html = personalized_html.replace(placeholder, str(row[column]))
                
                msg.attach(MIMEText(personalized_html, "html"))
                
                # Attach PDF if column exists
                if 'PDF_Path' in df.columns and pd.notna(row['PDF_Path']):
                    pdf_path = row['PDF_Path']
                    if os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as pdf_file:
                            pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
                            pdf_attachment.add_header(
                                'Content-Disposition',
                                'attachment',
                                filename=os.path.basename(pdf_path)
                            )
                            msg.attach(pdf_attachment)
                
                server.sendmail(smtp_email, str(row["Email"]), msg.as_string())
                sent_count += 1
                time.sleep(1)  # Rate limiting
            except Exception as e:
                continue
        
        server.quit()
        
        return jsonify({
            'status': 'success',
            'message': f'Successfully sent {sent_count} out of {total_count} emails',
            'sentCount': sent_count,
            'totalCount': total_count
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True)