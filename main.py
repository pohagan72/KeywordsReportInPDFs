import os
import time
import pandas as pd
import warnings
from azure.ai.formrecognizer import FormRecognizerClient
from azure.core.credentials import AzureKeyCredential
import re
import logging
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from PyPDF2 import PdfReader, PdfWriter
import io
from flask import Flask, request, render_template, send_file, jsonify
from threading import Thread, Lock
import zipfile

# Suppress warnings related to the openpyxl library
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Azure Form Recognizer credentials
AZURE_KEY = "KEY"  # Replace with your key
ENDPOINT = "ENDPOINT" # Replace with your endpoint

# Initialize logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize Azure client
try:
    azure_client = FormRecognizerClient(ENDPOINT, AzureKeyCredential(AZURE_KEY))
except Exception as e:
    logger.error(f"Failed to initialize Azure Form Recognizer Client: {e}")

# Flask app setup
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Global variables for progress tracking
total_images = 0
processed_images = 0
progress_lock = Lock()

def search_terms_in_lines(lines, search_terms):
    """
    Searches for exact matches of search terms in the lines and returns their coordinates.

    Args:
        lines (list): List of line objects from Form Recognizer.
        search_terms (list): List of terms to search for.

    Returns:
        list: List of dictionaries containing term, page number, and coordinates.
    """
    found_terms = []
    search_terms_lower = {term.lower(): term for term in search_terms}  # Map lowercase to original term for case preservation

    for line in lines:
        for word in line.words:
            word_text_lower = word.text.lower()
            if word_text_lower in search_terms_lower:
                found_terms.append({
                    "Keyword": search_terms_lower[word_text_lower],
                    "Page": line.page_number,
                    "Coordinates": word.bounding_box
                })
    return found_terms

def process_pdf_azure(pdf_path, search_terms, client):
    """
    Processes a PDF using Azure Form Recognizer to find pages containing specified search terms and their coordinates.

    Args:
        pdf_path (str): Path to the PDF file.
        search_terms (list): List of terms to search for.
        client (FormRecognizerClient): Initialized Azure Form Recognizer client.

    Returns:
        list: List of dictionaries with search results.
    """
    pdf_results = []
    try:
        with open(pdf_path, "rb") as f:
            poller = client.begin_recognize_content(f)
            result = poller.result()
    except Exception as e:
        logger.error(f"Error processing PDF with Azure: {e}")
        return pdf_results

    for page in result:
        found_terms = search_terms_in_lines(page.lines, search_terms)
        if found_terms:
            pdf_results.extend(found_terms)

    return pdf_results

def generate_excel_report(results, output_path):
    """
    Generates an Excel report from search results.

    Args:
        results (list): List of search result dictionaries.
        output_path (str): Path to save the Excel report.
    """
    if not results:
        logger.info("No keywords found.")
        return

    df = pd.DataFrame(results)
    try:
        df.to_excel(output_path, index=False)
        logger.info(f"Report saved to {output_path}")
    except Exception as e:
        logger.error(f"Error writing Excel report: {e}")

def extract_coordinates(coord_str):
    """
    Extracts coordinates from the "Point(x=..., y=...)" format.

    Args:
        coord_str (str): String containing coordinates.

    Returns:
        list: List of tuples (x, y) as floats.
    """
    points = re.findall(r"Point\(x=([\d.]+), y=([\d.]+)\)", coord_str)
    return [(float(x), float(y)) for x, y in points]

def highlight_pdf(pdf_path, excel_path, output_pdf_path):
    """
    Highlights keywords in a PDF based on coordinates from an Excel file.

    Args:
        pdf_path (str): Path to the PDF file.
        excel_path (str): Path to the Excel file with coordinates.
        output_pdf_path (str): Path to save the highlighted PDF.
    """
    df = pd.read_excel(excel_path)
    page_coordinates = {}

    for index, row in df.iterrows():
        keyword = row['Keyword']
        page = row['Page']
        coordinates = extract_coordinates(row['Coordinates'])

        if page not in page_coordinates:
            page_coordinates[page] = []

        page_coordinates[page].append((keyword, coordinates))

    pdf_reader = PdfReader(pdf_path)
    pdf_writer = PdfWriter()

    # Get the size of the first page
    pdf_page_size = pdf_reader.pages[0].mediabox
    pdf_width = float(pdf_page_size[2])
    pdf_height = float(pdf_page_size[3])

    for page_num in range(len(pdf_reader.pages)):
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(pdf_width, pdf_height))  # Use the extracted width and height
        c.setStrokeColorRGB(1, 0, 0)
        c.setLineWidth(1.5)

        if (page_num + 1) in page_coordinates:
            for keyword, points in page_coordinates[page_num + 1]:
                x_coords = [p[0] for p in points]
                y_coords = [p[1] for p in points]

                # Convert coordinates to points
                x_coords = [x * inch for x in x_coords]
                y_coords = [pdf_height - (y * inch) for y in y_coords]  # Adjust y-coordinate

                # Draw a rectangle around the keyword
                c.line(x_coords[0], y_coords[0], x_coords[1], y_coords[1])
                c.line(x_coords[1], y_coords[1], x_coords[2], y_coords[2])
                c.line(x_coords[2], y_coords[2], x_coords[3], y_coords[3])
                c.line(x_coords[3], y_coords[3], x_coords[0], y_coords[0])

        c.save()

        packet.seek(0)
        new_pdf = PdfReader(packet)
        page = pdf_reader.pages[page_num]
        page.merge_page(new_pdf.pages[0])
        pdf_writer.add_page(page)

    with open(output_pdf_path, "wb") as output_pdf:
        pdf_writer.write(output_pdf)

def process_folder(folder_path):
    global total_images, processed_images
    with progress_lock:
        processed_images = 0

    excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], "search_terms.xlsx")
    df_terms = pd.read_excel(excel_file_path)
    search_terms = df_terms.iloc[:, 0].dropna().tolist()

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    total_images = len(pdf_files)

    results = []
    highlighted_pdfs = [] # list of paths to the highlighted pdfs
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        pdf_name_no_ext = os.path.splitext(pdf_file)[0]
        output_excel_filename = os.path.join(folder_path, f"{pdf_name_no_ext}-analysis.xlsx")
        output_pdf_filename = os.path.join(folder_path, f"{pdf_name_no_ext}-highlighted.pdf")

        # Process PDF using Azure Form Recognizer
        pdf_results = process_pdf_azure(pdf_path, search_terms, azure_client)

        # Generate Excel report for the current PDF
        generate_excel_report(pdf_results, output_excel_filename)

        # Highlight PDF based on the generated Excel report
        if pdf_results:
            highlight_pdf(pdf_path, output_excel_filename, output_pdf_filename)
            highlighted_pdfs.append(output_pdf_filename)

        with progress_lock:
            processed_images += 1

        # Collect results for all files
        results.extend(pdf_results)

    # Create combined excel
    combined_excel_path = os.path.join(folder_path, "combined_analysis.xlsx")
    generate_excel_report(results, combined_excel_path)
    return highlighted_pdfs # return list of output pdfs

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    folder_path = request.form['folderPath']
    excel_file = request.files['excelFile']

    if not os.path.isdir(folder_path):
        return jsonify({'error': 'Invalid folder path.'}), 400

    if excel_file:
        excel_filename = "search_terms.xlsx"
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        excel_file.save(excel_file_path)
    else:
        return jsonify({'error': 'No Excel file uploaded.'}), 400

    # Start processing in a background thread
    thread = Thread(target=process_folder_async, args=[folder_path])
    thread.start()

    return jsonify({'message': 'Processing started successfully.'}), 200

def process_folder_async(folder_path):
    try:
        output_pdf_paths = process_folder(folder_path) # returns the paths to output pdfs

        # You might want to do something after processing, like sending a notification
        logger.info(f"Processing complete. Highlighted PDFs generated.")

    except Exception as e:
        logger.error(f"Error during processing: {e}")

@app.route('/progress')
def progress():
    with progress_lock:
        return jsonify({'total_images': total_images, 'processed_images': processed_images})

@app.route('/download')
def download():
    folder_path = request.args.get('folder_path')
    # zip all the pdfs in the folder path
    output_filename = os.path.join(folder_path, "highlighted_pdfs.zip")
    try:
        with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.endswith("-highlighted.pdf"):
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, os.path.relpath(file_path, folder_path))
    except Exception as e:
         return jsonify({'error': f"Error creating zip file: {e}"}), 500

    try:
        return send_file(output_filename, as_attachment=True, download_name="highlighted_pdfs.zip")
    except FileNotFoundError:
        return jsonify({'error': 'File not found. Processing might not be complete.'}), 404
    except Exception as e:
        return jsonify({'error': f"Error sending file: {e}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
