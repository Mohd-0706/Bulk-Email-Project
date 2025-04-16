# Bulk Email Sender Website

A colorful, animated web application for sending bulk emails with Excel integration.

## Features
- Drag-and-drop merge tags from Excel headers
- SMTP email sending with Gmail/Yahoo/Outlook
- Real-time status notifications
- Responsive design
- Secure file uploads

## Requirements
- Python 3.6+
- Node.js (for frontend development)
- Excel file with recipient data


2. Run the backend server:
```bash
python app.py
```

3. Open `index.html` in a web browser

## Usage
1. Enter your email and app password
2. Upload an Excel file with recipient data
3. Compose your message using merge tags
4. Click "Send Emails"

## SMTP Setup Instructions
1. **Gmail**:
   - Enable "Less secure app access" in Google Account settings
   - Or create an App Password if using 2FA

2. **Yahoo/Outlook**:
   - Use your normal password
   - May need to enable "Allow less secure apps"

## Notes
- The application runs on `http://localhost:5000`
- Uploaded files are stored in the `uploads` folder
- Email sending runs in background threads
