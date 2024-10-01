from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename
import os
from pdf2docx import Converter
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'  # Folder to save uploaded files
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert_pdf_to_word', methods=['POST'])
def convert_pdf_to_word():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    # Save the uploaded PDF file
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
    file.save(pdf_path)

    # Convert PDF to Word
    docx_path = pdf_path.replace('.pdf', '.docx')
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

    # Send the converted file back
    return send_from_directory(app.config['UPLOAD_FOLDER'], os.path.basename(docx_path), as_attachment=True)

@app.route('/split_pdf', methods=['POST'])
def split_pdf():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    # Save the uploaded PDF file
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
    file.save(pdf_path)

    # Split PDF logic (assuming we want to split by pages)
    pdf_reader = PdfReader(pdf_path)
    pdf_writer = PdfWriter()

    # Example: Split each page into a separate PDF
    for i in range(len(pdf_reader.pages)):
        pdf_writer.add_page(pdf_reader.pages[i])
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'split_page_{i + 1}.pdf')
        with open(output_path, 'wb') as out_file:
            pdf_writer.write(out_file)
        pdf_writer = PdfWriter()  # Reset the writer for the next page

    return "PDF split into individual pages."

@app.route('/merge_pdf', methods=['POST'])
def merge_pdf():
    if 'files' not in request.files:
        return "No file part", 400
    files = request.files.getlist('files')
    pdf_writer = PdfWriter()

    # Save each uploaded PDF file and add to writer
    for file in files:
        if file.filename == '':
            return "No selected file", 400
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(pdf_path)
        pdf_reader = PdfReader(pdf_path)
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

    merged_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], 'merged.pdf')
    with open(merged_pdf_path, 'wb') as out_file:
        pdf_writer.write(out_file)

    return send_from_directory(app.config['UPLOAD_FOLDER'], 'merged.pdf', as_attachment=True)

@app.route('/convert_pdf_to_image', methods=['POST'])
def convert_pdf_to_image():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    # Save the uploaded PDF file
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
    file.save(pdf_path)

    # Convert PDF to images
    images = convert_from_path(pdf_path)
    image_paths = []

    for i, image in enumerate(images):
        image_path = os.path.join(app.config['UPLOAD_FOLDER'], f'page_{i + 1}.jpg')
        image.save(image_path, 'JPEG')
        image_paths.append(image_path)

    return "PDF converted to images."

@app.route('/convert_image_to_pdf', methods=['POST'])
def convert_image_to_pdf():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    # Save the uploaded image file
    image_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
    file.save(image_path)

    # Convert image to PDF
    pdf_path = image_path.replace('.jpg', '.pdf').replace('.png', '.pdf')
    image = Image.open(image_path)
    image.save(pdf_path, "PDF", resolution=100.0)

    return send_from_directory(app.config['UPLOAD_FOLDER'], os.path.basename(pdf_path), as_attachment=True)

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)

if __name__ == '__main__':
    app.run(debug=True)
