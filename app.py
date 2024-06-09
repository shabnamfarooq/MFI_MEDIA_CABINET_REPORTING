from flask import Flask, request, redirect, url_for, send_file, render_template, flash
import pandas as pd
import os
import openpyxl
import logging

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Needed for flash messages

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(RESULT_FOLDER):
    os.makedirs(RESULT_FOLDER)

def compare_reports(members_file, report_file, output_file):
    try:
        members_df = pd.read_excel(members_file, sheet_name='Members')
        reports_df = pd.read_excel(report_file, sheet_name='Reports')

        members_df['Name'] = members_df['Name'].str.strip().str.lower()
        reports_df['Name'] = reports_df['Name'].str.strip().str.lower()

        members_df['Status'] = members_df['Name'].apply(
            lambda x: 'Submitted' if x in reports_df['Name'].values else 'Not Submitted')

        members_df.to_excel(output_file, index=False)
    except Exception as e:
        logging.error(f"Error comparing reports: {e}")
        raise

@app.route('/')
def upload_form():
    try:
        return render_template('upload.html')
    except Exception as e:
        logging.error(f"Error rendering upload form: {e}")
        return "An error occurred while rendering the form.", 500

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'members' not in request.files or 'report' not in request.files:
        flash('No file part')
        return redirect(request.url)

    members_file = request.files['members']
    report_file = request.files['report']

    if members_file.filename == '' or report_file.filename == '':
        flash('No selected file')
        return redirect(request.url)

    members_path = os.path.join(app.config['UPLOAD_FOLDER'], 'members.xlsx')
    report_path = os.path.join(app.config['UPLOAD_FOLDER'], 'report.xlsx')
    result_path = os.path.join(app.config['RESULT_FOLDER'], 'result.xlsx')

    try:
        members_file.save(members_path)
        report_file.save(report_path)

        compare_reports(members_path, report_path, result_path)

        return send_file(result_path, as_attachment=True)
    except Exception as e:
        logging.error(f"Error uploading files: {e}")
        flash('An error occurred while processing the files.')
        return redirect(request.url)

if __name__ == "__main__":
    from waitress import serve
    port = int(os.environ.get('PORT', 8000))
    serve(app, host='0.0.0.0', port=port)
