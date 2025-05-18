import os
import uuid
from flask import Flask, request, send_file, render_template_string
from pdf2docx import Converter

app = Flask(__name__)

# Get the base directory dynamically
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
CONVERTED_FOLDER = os.path.join(BASE_DIR, 'converted')

# Create the folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

# Simple HTML form
HTML_FORM = '''
<!DOCTYPE html>
<html>
<head><title>PDF to Word</title></head>
<body>
    <h2>Convert PDF to Word</h2>
    <form method="POST" action="/convert" enctype="multipart/form-data">
        <input type="file" name="pdf_file" accept="application/pdf" required><br><br>
        <input type="submit" value="Convert">
    </form>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_FORM)

@app.route('/convert', methods=['POST'])
def convert():
    if 'pdf_file' not in request.files:
        return 'No file uploaded.'

    file = request.files['pdf_file']
    if file.filename == '':
        return 'No file selected.'

    # Save the uploaded PDF
    pdf_filename = str(uuid.uuid4()) + '.pdf'
    pdf_path = os.path.join(UPLOAD_FOLDER, pdf_filename)
    file.save(pdf_path)

    # Generate the output Word file path
    docx_filename = str(uuid.uuid4()) + '.docx'
    docx_path = os.path.join(CONVERTED_FOLDER, docx_filename)

    # Convert PDF to DOCX
    converter = Converter(pdf_path)
    converter.convert(docx_path)
    converter.close()

    # Ensure file is created before sending
    if not os.path.exists(docx_path):
        return 'Conversion failed. DOCX file not found.'

    return send_file(docx_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
