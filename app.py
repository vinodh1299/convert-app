from flask import Flask, request, render_template
import os
import tempfile
import zipfile
from pdf2image import convert_from_path
from docx import Document
from fpdf import FPDF
from werkzeug.utils import secure_filename

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    # Get form data
    conversion_type = request.form['conversion_type']
    output_folder = request.form.get('output_folder', tempfile.gettempdir())

    print(f"Conversion type selected: {conversion_type}")
    print(f"Output folder: {output_folder}")

    # Ensure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Output folder created: {output_folder}")

    # Handle ZIP file
    input_zip = request.files.get('input_zip')
    if input_zip and input_zip.filename.endswith('.zip'):
        temp_dir = tempfile.gettempdir()
        zip_filename = secure_filename(input_zip.filename)
        zip_path = os.path.join(temp_dir, zip_filename)
        input_zip.save(zip_path)

        print(f"ZIP file saved at: {zip_path}")

        # Extract ZIP contents
        input_folder = os.path.join(temp_dir, 'input_folder')
        os.makedirs(input_folder, exist_ok=True)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(input_folder)
            print(f"Extracted files: {zip_ref.namelist()}")

        converted_files = []
        for filename in os.listdir(input_folder):
            file_path = os.path.join(input_folder, filename)
            if os.path.isfile(file_path):
                print(f"Processing file: {file_path}")

                try:
                    if conversion_type == "pdf_to_word":
                        print("Calling convert_pdf_to_word")
                        convert_pdf_to_word(file_path, output_folder)
                    elif conversion_type == "word_to_pdf":
                        print("Calling convert_word_to_pdf")
                        convert_word_to_pdf(file_path, output_folder)
                    elif conversion_type == "image_to_pdf":
                        print("Calling convert_image_to_pdf")
                        convert_image_to_pdf(file_path, output_folder)
                    elif conversion_type == "pdf_to_image":
                        print("Calling convert_pdf_to_image")
                        convert_pdf_to_image(file_path, output_folder)

                    # Collect converted files
                    for converted_file in os.listdir(output_folder):
                        if filename.split('.')[0] in converted_file:
                            converted_files.append(converted_file)
                            print(f"Converted file: {converted_file}")

                except Exception as e:
                    print(f"Error during conversion: {str(e)}")
                    return render_template('completion.html', message=f"Error converting file: {file_path}", files=[])

        print(f"Converted files: {converted_files}")
        if converted_files:
            return render_template('completion.html', message="Conversion complete!", files=converted_files)
        else:
            return render_template('completion.html', message="No valid files were converted.", files=[])

    else:
        print("Invalid ZIP file.")
        return "Invalid ZIP file. Please upload a valid ZIP file."

# Conversion Functions
def convert_pdf_to_word(pdf_path, output_folder):
    print(f"Converting PDF to Word: {pdf_path}")
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    word_path = os.path.join(output_folder, f'{pdf_name}.docx')
    doc = Document()
    doc.add_paragraph("This is a placeholder text extracted from PDF.")
    doc.save(word_path)
    print(f'Saved {word_path}')

def convert_word_to_pdf(word_path, output_folder):
    print(f"Converting Word to PDF: {word_path}")
    pdf_name = os.path.splitext(os.path.basename(word_path))[0]
    pdf_path = os.path.join(output_folder, f'{pdf_name}.pdf')
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="This is a placeholder PDF generated from Word.", ln=True)
    pdf.output(pdf_path)
    print(f'Saved {pdf_path}')

def convert_image_to_pdf(image_path, output_folder):
    print(f"Converting Image to PDF: {image_path}")
    pdf_name = os.path.splitext(os.path.basename(image_path))[0]
    pdf_path = os.path.join(output_folder, f'{pdf_name}.pdf')
    pdf = FPDF()
    pdf.add_page()
    pdf.image(image_path, x=10, y=10, w=180)
    pdf.output(pdf_path)
    print(f'Saved {pdf_path}')

def convert_pdf_to_image(pdf_path, output_folder):
    print(f"Converting PDF to Image: {pdf_path}")
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    images = convert_from_path(pdf_path)
    
    for i, image in enumerate(images):
        image_path = os.path.join(output_folder, f'{pdf_name}_page_{i + 1}.jpg')
        image.save(image_path, 'JPEG')
        print(f'Saved {image_path}')

if __name__ == "__main__":
    app.run(debug=True)
