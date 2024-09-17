from flask import Flask, request, jsonify, send_file, render_template
import os
import zipfile
import shutil
from werkzeug.utils import secure_filename
from pdf2image import convert_from_path
from fpdf import FPDF
import subprocess
import win32com.client  # For Windows users only
import pythoncom
from pdf2docx import Converter  # You need to install pdf2docx library
import time

app = Flask(__name__)

# Upload folder setup
UPLOAD_FOLDER = './uploads'
OUTPUT_FOLDER = './output'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# Helper function to extract base filename (without folder and extension)
def get_base_filename(filepath):
    return os.path.splitext(os.path.basename(filepath))[0]

# Route for serving the index.html (Homepage)
@app.route('/')
def home():
    return render_template('index.html')

# Route to handle file conversion
@app.route('/convert', methods=['POST'])
def convert_files():
    files = request.files.getlist('files')
    conversion_type = request.form['conversionType']
    
    if not files:
        return jsonify({'error': 'No files selected'}), 400
    
    output_files = []

    for file in files:
        base_filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], base_filename)
        file.save(filepath)

        # Print the path to ensure it's correct
        print ('File saved at:', filepath)

        try:
            if conversion_type == 'pdf_to_word':
                output_files.append(convert_pdf_to_word(filepath))
            elif conversion_type == 'word_to_pdf':
                output_files.append(convert_word_to_pdf(filepath))
            elif conversion_type == 'pdf_to_image':
                output_files.extend(convert_pdf_to_image(filepath))
            elif conversion_type == 'image_to_pdf':
                output_files.append(convert_image_to_pdf(filepath))
            else:
                return jsonify({'error': 'Invalid conversion type'}), 400
        except Exception, e:
            return jsonify({'error': str(e)}), 500

    # Check if there are any output files to zip
    if not output_files:
        return jsonify({'error': 'No files were converted successfully'}), 500

    zip_path = zip_files(output_files)
    
    return jsonify({'downloadLink': '/download/' + zip_path})

# Conversion function to handle PDF to Word conversion
def convert_pdf_to_word(filepath):
    base_name = get_base_filename(filepath)
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], base_name + '.docx')
    
    try:
        cv = Converter(filepath)
        cv.convert(output_path, start=0, end=None)
        cv.close()
    except Exception, e:
        print ('Error converting PDF to Word:', e)
    
    return output_path

# Conversion function to handle Word to PDF conversion
def convert_word_to_pdf(filepath):
    base_name = get_base_filename(filepath)
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], base_name + '.pdf')

    pythoncom.CoInitialize()
    
    try:
        if os.name == 'nt':  # For Windows
            word = win32com.client.Dispatch('Word.Application')
            
            # Wait for a short period to ensure file system operations are completed
            time.sleep(1)
            
            # Get absolute path for both input and output
            abs_filepath = os.path.abspath(filepath)
            abs_output_path = os.path.abspath(output_path)
            
            try:
                # Verify the file path
                print ('Attempting to open file:', abs_filepath)
                
                doc = word.Documents.Open(abs_filepath)
                doc.SaveAs(abs_output_path, FileFormat=17)  # 17 is the format for PDF
                doc.Close()
            except Exception, e:
                print ('Error opening or saving file:', e)
                raise
            finally:
                word.Quit()
        else:
            command = [
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', app.config['OUTPUT_FOLDER'], filepath
            ]
            subprocess.call(command)
    
    except Exception, e:
        print ('Error in convert_word_to_pdf:', e)
        raise
    finally:
        pythoncom.CoUninitialize()
    
    return output_path


# Conversion function to handle PDF to Image conversion
def convert_pdf_to_image(filepath):
    try:
        images = convert_from_path(filepath)
    except Exception, e:
        print ('Error converting PDF to images:', e)
        return []

    output_paths = []
    base_name = get_base_filename(filepath)

    for i, image in enumerate(images):
        if len(images) == 1:
            output_filename = base_name + '.jpg'
        else:
            output_filename = base_name + '.jpg'
        
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        try:
            image.save(output_path, 'JPEG')
            output_paths.append(output_path)
        except Exception, e:
            print ('Error saving image', output_filename, ':', e)

    return output_paths

# Conversion function to handle Image to PDF conversion
def convert_image_to_pdf(filepath):
    base_name = get_base_filename(filepath)
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], base_name + '.pdf')
    
    pdf = FPDF()
    pdf.add_page()
    try:
        pdf.image(filepath, x=10, y=10, w=190)
        pdf.output(output_path, 'F ')
    except Exception, e:
        print ('Error creating PDF from image:', e)
    
    return output_path

# Helper function to zip the converted files
def zip_files(file_paths):
    zip_filename = 'converted_files.zip'
    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file_path in file_paths:
            if os.path.isfile(file_path):  # Check if file exists
                zipf.write(file_path, os.path.basename(file_path))
            else:
                print ('File not found for zipping:', file_path)
    
    return zip_filename

# Helper function to clear the upload and output folders
def clear_folders():
    for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception, e:
                print ('Failed to delete', file_path, '. Reason:', e)

# Route to serve the zipped files and clear folders afterward
@app.route('/download/<zip_filename>')
def download(zip_filename):
    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
    
    if not os.path.isfile(zip_path):
        return jsonify({'error': 'File not found'}), 404

    try:
        response = send_file(zip_path, as_attachment=True)
    except Exception, e:
        print ('Error sending file:', e)
        return jsonify({'error': 'Error sending file'}), 500
    
    # Clear the upload and output folders after the file is sent
    clear_folders()
    
    return response

if __name__ == '__main__':
    app.run(debug=True)