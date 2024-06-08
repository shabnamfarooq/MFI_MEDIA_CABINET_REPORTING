from flask import Flask, request, redirect, url_for, send_file, render_template
import pandas as pd
import os
import openpyxl

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(RESULT_FOLDER):
    os.makedirs(RESULT_FOLDER)


def compare_reports(members_file, report_file, output_file):
    members_df = pd.read_excel(members_file, sheet_name='Members')
    reports_df = pd.read_excel(report_file, sheet_name='Reports')

    members_df['Name'] = members_df['Name'].str.strip().str.lower()
    reports_df['Name'] = reports_df['Name'].str.strip().str.lower()

    members_df['Status'] = members_df['Name'].apply(
        lambda x: 'Submitted' if x in reports_df['Name'].values else 'Not Submitted')

    members_df.to_excel(output_file, index=False)


@app.route('/')
def upload_form():
    return render_template('upload.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'members' not in request.files or 'report' not in request.files:
        return redirect(request.url)

    members_file = request.files['members']
    report_file = request.files['report']

    if members_file.filename == '' or report_file.filename == '':
        return redirect(request.url)

    members_path = os.path.join(app.config['UPLOAD_FOLDER'], 'members.xlsx')
    report_path = os.path.join(app.config['UPLOAD_FOLDER'], 'report.xlsx')
    result_path = os.path.join(app.config['RESULT_FOLDER'], 'result.xlsx')

    members_file.save(members_path)
    report_file.save(report_path)

    compare_reports(members_path, report_path, result_path)

    return send_file(result_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
