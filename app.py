from flask import Flask, render_template, request, jsonify
import smtplib
import time
import pandas as pd
import numpy as np
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import json

app = Flask(__name__)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Custom JSON encoder to handle NaN values
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
        df = pd.read_excel(excel_file)
        
        # Convert DataFrame to dict with proper NaN handling
        preview_data = []
        for _, row in df.head(5).iterrows():
            row_dict = {}
            for col in df.columns:
                value = row[col]
                if pd.isna(value):
                    row_dict[col] = None
                else:
                    row_dict[col] = value
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
        # Get form data
        smtp_email = request.form['smtp_email']
        smtp_password = request.form['smtp_password']
        subject_template = request.form['subject']
        html_template = request.form['template']
        
        # Get the Excel file
        excel_file = request.files['excel_file']
        df = pd.read_excel(excel_file)
        
        # Replace NaN values with empty strings for email templates
        df = df.fillna('')
        
        sent_count = 0
        total_count = len(df)
        
        # Establish SMTP connection
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(smtp_email, smtp_password)
        
        for index, row in df.iterrows():
            try:
                # Create personalized email
                msg = MIMEMultipart()
                msg["From"] = smtp_email
                msg["To"] = str(row["Email"])
                msg["Subject"] = subject_template
                
                # Personalize template
                personalized_html = html_template
                for column in df.columns:
                    placeholder = f"{{{column}}}"
                    if placeholder in html_template:
                        personalized_html = personalized_html.replace(
                            placeholder, str(row[column]))
                
                msg.attach(MIMEText(personalized_html, "html"))
                
                # Send email
                server.sendmail(smtp_email, str(row["Email"]), msg.as_string())
                sent_count += 1
                time.sleep(1)  # Rate limiting
            except Exception as e:
                print(f"Error sending to {row['Email']}: {str(e)}")
                continue
        
        server.quit()
        
        return jsonify({
            'status': 'success',
            'message': f'Successfully sent {sent_count} out of {total_count} emails',
            'sentCount': sent_count,
            'totalCount': total_count
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        })

if __name__ == '__main__':
    app.run(debug=True)