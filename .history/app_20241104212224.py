import os
import pandas as pd
from flask import Flask, request, render_template, send_file, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULT_FOLDER'] = 'results'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part in the request"
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    # Simpan file ke folder uploads
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    # Baca file Excel dan lakukan pencocokan data
    excel_data = pd.ExcelFile(file_path)
    sheet1_df = excel_data.parse('Sheet1')
    sheet2_df = excel_data.parse('Sheet2')

    if 'sia' in sheet1_df.columns and 'sia' in sheet2_df.columns:
        # Left join dari Sheet2 ke Sheet1 untuk mencocokkan semua data 'sia' dari Sheet2
        matching_sia = pd.merge(sheet2_df, sheet1_df, on='sia', how='left')

        # Simpan hasil ke file baru di folder results
        result_filename = f"result_{filename}"
        result_file_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)
        with pd.ExcelWriter(result_file_path, engine='openpyxl') as writer:
            excel_data.parse('Sheet1').to_excel(writer, sheet_name='Sheet1', index=False)
            excel_data.parse('Sheet2').to_excel(writer, sheet_name='Sheet2', index=False)
            matching_sia.to_excel(writer, sheet_name='Sheet3', index=False)

        # Redirect ke halaman download dengan filename hasil
        return redirect(url_for('download_page', filename=result_filename))

    return "Kolom 'sia' tidak ditemukan di salah satu atau kedua sheet."

@app.route('/download/<filename>')
def download_page(filename):
    # Halaman untuk menampilkan link unduhan hasil
    return render_template('download.html', filename=filename)

@app.route('/download_file/<filename>')
def download_file(filename):
    # Kirim file hasil untuk diunduh
    result_file_path = os.path.join(app.config['RESULT_FOLDER'], filename)
    return send_file(result_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
