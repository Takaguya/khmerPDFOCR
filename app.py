import os
import pytesseract
import cv2
import numpy as np
from flask import Flask, request, render_template, send_from_directory, redirect, url_for
from docx import Document
from docx.shared import Pt
from pdf2image import convert_from_path
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='statics')

# Configuration for directories (relative to the location of the Flask app)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # This gets the folder where the script is located
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')  # Relative path for upload folder
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'document_output')  # Relative path for output folder

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Ensure upload and output folders exist (created in the same directory as the app)
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def process_image(image_path, doc):
    """Process a single image file and append OCR results to a .docx file."""
    image = cv2.imread(image_path)
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    resized_image = cv2.resize(gray_image, (int(gray_image.shape[1] * 1.75), int(gray_image.shape[0] * 1.75)))
    text = pytesseract.image_to_string(resized_image, config=r'-l khm+eng --psm 3')
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    run.font.name = 'Khmer OS Battambong'
    run.font.size = Pt(13)

def process_pdf(pdf_file, output_file_path):
    """Process a PDF file and save OCR results to a combined .docx file."""
    doc = Document()
    images = convert_from_path(pdf_file)
    for image in images:
        image_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray_image = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
        resized_image = cv2.resize(image_cv, (int(gray_image.shape[1] * 1.75), int(gray_image.shape[0] * 1.75)))
        text = pytesseract.image_to_string(resized_image, config=r'-l khm+eng --psm 3')
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(text)
        run.font.name = 'Khmer OS Battambong'
        run.font.size = Pt(13)
    doc.save(output_file_path)

@app.route('/')
def index():
    error_message = request.args.get('error')  # Get the error message from the URL, if any
    return render_template('index.html', error_message=error_message)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return redirect(url_for('index', error="No file uploaded."))

    file = request.files['file']
    if file.filename == '':
        return redirect(url_for('index', error="No file selected."))

    # Save the uploaded file
    filename = secure_filename(file.filename)
    upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(upload_path)

    # Determine output file path
    output_filename = f"{os.path.splitext(filename)[0]}.docx"
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

    # Determine file type and process accordingly
    file_ext = os.path.splitext(filename)[1].lower()
    if file_ext == '.pdf':
        process_pdf(upload_path, output_path)
    elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
        doc = Document()
        process_image(upload_path, doc)
        doc.save(output_path)
    else:
        # Return an error message if file type is not supported
        return redirect(url_for('index', error="Unsupported file type. Please upload a PDF or an image file."))

    return redirect(url_for('download', filename=output_filename))

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
