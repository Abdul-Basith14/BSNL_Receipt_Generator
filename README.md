# BSNL Cash Receipts Generator

A web application for generating formatted cash receipts from Excel data.

## Features

- 📤 Upload Excel files with work details
- ⚡ Automatic receipt generation
- 📥 Download generated receipts
- 🎨 Modern, user-friendly interface
- 🚀 Deployable to cloud platforms

## Installation

1. Install Python 3.8 or higher

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Running Locally

```bash
python app.py
```

Open your browser to http://localhost:5000

## File Format

Your Excel file should contain columns similar to TY Adv Appl format:
- **Column 1**: Date
- **Column 2**: Route
- **Column 3**: Work Details
- **Column 7**: Pits/OH Cable indicator
- **Column 8**: Amount

## Deployment Options

### 1. Deploy to Render.com (Recommended)

1. Create account at https://render.com
2. Create new Web Service
3. Connect your GitHub repository
4. Set build command: `pip install -r requirements.txt`
5. Set start command: `python app.py`
6. Deploy!

### 2. Deploy to Heroku

1. Install Heroku CLI
2. Create Procfile:
```
web: python app.py
```
3. Deploy:
```bash
heroku login
heroku create bsnl-receipts
git push heroku main
```

### 3. Deploy to PythonAnywhere

1. Upload files to PythonAnywhere
2. Create web app with Flask
3. Set WSGI configuration
4. Reload web app

### 4. Deploy to Railway.app

1. Connect GitHub repository
2. Railway auto-detects Flask app
3. Deploy with one click

## Technical Details

- **Backend**: Flask (Python web framework)
- **Excel Processing**: openpyxl library
- **File Handling**: In-memory processing (no disk storage)
- **Max File Size**: 16MB
- **Supported Formats**: .xlsx, .xls

## Security Notes

- Files are processed in memory
- No permanent storage of uploads
- HTTPS recommended for production
- Set strong SECRET_KEY in production

## License

© 2026 BSNL Cash Receipts Generator
