# -*- coding: utf-8 -*-
import os # <--- Make sure os is imported
import io
import zipfile
import logging
import pandas as pd
import re
import traceback
import time
import uuid # For generating unique blob names

from flask import Flask, render_template, request, send_file, jsonify, abort
from azure.ai.formrecognizer import FormRecognizerClient
from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import HttpResponseError, ServiceRequestError, ResourceNotFoundError
from azure.core.polling import LROPoller # Optional for type hints
# --- Azure Storage ---
# Keep BlobServiceClient for potential future use or if other parts need it,
# but we won't use it for the intermediate steps in process_files
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient
# ---------------------
# --- Using pypdf ---
from pypdf import PdfReader, PdfWriter
# -----------------------
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from werkzeug.utils import secure_filename
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

# -------------------- Flask App Setup --------------------
app = Flask(__name__, static_folder='static')
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024 # 500 MB limit

# Rate Limiting
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["200 per day", "50 per hour"]
)

# -------------------- Logging Setup ----------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(name)s - Thread: %(thread)d - %(message)s'
)
logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
logging.getLogger("azure.storage.blob").setLevel(logging.WARNING)
logging.getLogger("pypdf").setLevel(logging.WARNING) # Quiet pypdf logs unless error
logger = logging.getLogger(__name__)

# -------------------- Azure Form Recognizer Setup --------
# WARNING: Hardcoded key! Acceptable ONLY for local testing as requested.
AZURE_FORM_RECOGNIZER_KEY = "KEY"
AZURE_FORM_RECOGNIZER_ENDPOINT = "Endpoint"

try:
    azure_form_recognizer_client = FormRecognizerClient(
        AZURE_FORM_RECOGNIZER_ENDPOINT,
        AzureKeyCredential(AZURE_FORM_RECOGNIZER_KEY),
        logging_enable=False
    )
    logger.info("Azure Form Recognizer client initialized successfully.")
except Exception as e:
    logger.error(f"FATAL: Failed to initialize Azure Form Recognizer client: {e}", exc_info=True)
    azure_form_recognizer_client = None

AZURE_STORAGE_CONNECTION_STRING = "Connection_String"
AZURE_STORAGE_CONTAINER_NAME = "tempfiles" # Still useful for potential archival/debug

try:
    if not AZURE_STORAGE_CONNECTION_STRING:
         raise ValueError("Azure Storage Connection String is not set.")
    # Initialize client but be aware it's not used in the main processing flow anymore
    azure_blob_service_client = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
    logger.info("Azure Blob Storage client initialized (may not be used in critical path).")
    # Optional: Check container access if needed elsewhere or for optional archival
    # try:
    #     container_client = azure_blob_service_client.get_container_client(AZURE_STORAGE_CONTAINER_NAME)
    #     container_client.get_container_properties()
    #     logger.info(f"Successfully accessed storage container '{AZURE_STORAGE_CONTAINER_NAME}'.")
    # except Exception as container_ex:
    #      logger.warning(f"Could not access storage container '{AZURE_STORAGE_CONTAINER_NAME}': {container_ex}")
    #      # Don't make this fatal if primary flow doesn't depend on it
except Exception as e:
    logger.error(f"Failed to initialize Azure Storage client: {e}", exc_info=True)
    azure_blob_service_client = None # Mark as unavailable

# -------------------- Utility Functions (Signatures Updated) --------------------

def read_search_terms(excel_stream: io.BytesIO):
    """Reads search terms directly from an Excel BytesIO stream."""
    logger.info(f"Reading search terms from provided Excel stream.")
    try:
        excel_stream.seek(0) # Ensure stream is at the beginning
        df = pd.read_excel(excel_stream, usecols=[0], header=None)
        search_terms_original = df.iloc[:, 0].dropna().astype(str).unique().tolist()
        if not search_terms_original:
            raise ValueError("No search terms found in Excel file.")
        search_terms_lower_set = set()
        lower_to_original_map = {}
        for term in search_terms_original:
            lower_term = term.lower()
            search_terms_lower_set.add(lower_term)
            if lower_term not in lower_to_original_map:
                 lower_to_original_map[lower_term] = term
        logger.info(f"Read {len(search_terms_original)} unique search terms from stream.")
        return search_terms_original, search_terms_lower_set, lower_to_original_map
    except (ValueError) as e:
         logger.error(f"Error reading Excel stream: {e}", exc_info=False)
         raise ValueError(f"Error reading or processing Excel file from stream: {e}")
    except Exception as e:
        logger.error(f"Unexpected error reading Excel stream: {e}", exc_info=True)
        raise RuntimeError(f"Unexpected error reading Excel file from stream: {e}")

def generate_excel_report(results):
    """Generates an Excel report (returns BytesIO stream)."""
    if not results:
        df = pd.DataFrame(columns=["Keyword", "Page", "X1", "Y1", "X2", "Y2", "X3", "Y3", "X4", "Y4"])
    else:
        formatted_results = []
        for item in results:
            coords = item['WordObject'].bounding_box
            if len(coords) == 4:
                 formatted_results.append({
                    "Keyword": item['Keyword'], "Page": item['Page'],
                    "X1": coords[0].x, "Y1": coords[0].y, "X2": coords[1].x, "Y2": coords[1].y,
                    "X3": coords[2].x, "Y3": coords[2].y, "X4": coords[3].x, "Y4": coords[3].y,
                })
            else: logger.warning(f"Keyword '{item['Keyword']}' page {item['Page']} bad bbox len: {len(coords)}. Skipping.")
        if not formatted_results: df = pd.DataFrame(columns=["Keyword", "Page", "X1", "Y1", "X2", "Y2", "X3", "Y3", "X4", "Y4"])
        else: df = pd.DataFrame(formatted_results)
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer: df.to_excel(writer, index=False, sheet_name="Analysis")
        output.seek(0)
        logger.info("Excel report generated successfully into memory stream.")
        return output
    except Exception as e:
        logger.error(f"Error generating Excel report: {e}", exc_info=True)
        raise RuntimeError(f"Error generating Excel report: {e}")

def process_pdf_azure(pdf_stream: io.BytesIO, search_terms_lower_set, lower_to_original_map, polling_interval_seconds=20):
    """Processes a PDF directly from a BytesIO stream using Form Recognizer."""
    if not azure_form_recognizer_client: raise ConnectionError("Azure Form Recognizer client not available.")
    try:
        logger.info(f"Starting Azure Form Recognizer analysis on provided PDF stream.")
        pdf_stream.seek(0) # Ensure stream is at the beginning for FR
        poller = azure_form_recognizer_client.begin_recognize_content(pdf_stream, content_type='application/pdf', logging_enable=False)
        logger.info(f"Polling Azure for results every {polling_interval_seconds} seconds...")
        start_poll_time = time.time()
        while not poller.done():
            try:
                current_status = poller.status()
                logger.debug(f"Polling... Current status: {current_status}")
                if poller.done(): break
                poller.wait(polling_interval_seconds)
            except TimeoutError: logger.debug(f"Polling timeout, checking status again..."); continue
            except (HttpResponseError, ServiceRequestError) as poll_err:
                logger.error(f"Error during Azure polling: {poll_err}", exc_info=True)
                raise ConnectionError(f"Polling Azure status failed: {poll_err}") from poll_err
            except Exception as poll_err:
                 logger.error(f"Unexpected error during Azure polling wait: {poll_err}", exc_info=True)
                 raise RuntimeError(f"Polling failed unexpectedly: {poll_err}") from poll_err
        logger.info(f"Azure polling loop completed in {time.time() - start_poll_time:.2f} seconds.")
        final_status = poller.status()
        logger.info(f"Azure job finished with status: {final_status}")
        if final_status.lower() == "succeeded": result = poller.result()
        elif final_status.lower() == "failed":
             try: poller.result() # Should raise
             except HttpResponseError as e: raise ConnectionError(f"Azure job failed ({e.status_code}).") from e
             except Exception as e: raise RuntimeError(f"Azure job failed '{final_status}': {e}") from e
             raise RuntimeError(f"Azure job failed '{final_status}', no error raised.")
        else: raise RuntimeError(f"Azure job unexpected status: {final_status}")
        pdf_results = []
        non_word_char_regex = re.compile(r'[^\w\s]')
        for page in result:
            if page.lines:
                for line in page.lines:
                    for word in line.words:
                        cleaned_word = non_word_char_regex.sub('', word.text).lower()
                        if cleaned_word in search_terms_lower_set:
                            original_term = lower_to_original_map.get(cleaned_word, word.text)
                            pdf_results.append({"Keyword": original_term, "Page": page.page_number, "WordObject": word})
        logger.info(f"Found {len(pdf_results)} keyword instances via Azure from PDF stream.")
        return pdf_results
    except (ConnectionError, RuntimeError) as e: raise e
    except Exception as e: raise RuntimeError(f"Unexpected error during Azure processing: {e}") from e

def highlight_pdf(original_pdf_stream: io.BytesIO, results_data):
    """
    Highlights keywords in a PDF provided as a BytesIO stream using pypdf.
    Returns the highlighted PDF content as a BytesIO stream.
    """
    if not results_data:
        logger.warning("No results data provided, skipping highlighting.")
        original_pdf_stream.seek(0)
        return io.BytesIO(original_pdf_stream.read())

    page_coords = {}
    for item in results_data:
        page_idx = item['Page']
        word_obj = item['WordObject']
        coords_raw = word_obj.bounding_box
        if len(coords_raw) == 4:
            try: coords = [(float(p.x), float(p.y)) for p in coords_raw]
            except (ValueError, TypeError) as coord_err:
                logger.warning(f"Invalid coords word '{item['Keyword']}' page {page_idx}. Skip. Err: {coord_err}")
                continue
        else:
            logger.warning(f"Word '{item['Keyword']}' page {page_idx} bad bbox len: {len(coords_raw)}. Skip.")
            continue
        page_map_idx = page_idx
        if page_map_idx not in page_coords: page_coords[page_map_idx] = []
        page_coords[page_map_idx].append(coords)

    output_pdf = io.BytesIO()
    try:
        original_pdf_stream.seek(0)
        original_pdf_reader = PdfReader(original_pdf_stream)
        pdf_writer = PdfWriter()
        num_pages = len(original_pdf_reader.pages)
        logger.info(f"Processing {num_pages} pages from PDF stream for highlighting (add_blank strategy).")

        for page_num in range(num_pages):
            page_index_1_based = page_num + 1
            original_page = original_pdf_reader.pages[page_num]

            try:
                page_box = original_page.mediabox
                pdf_width, pdf_height = float(page_box.width), float(page_box.height)
                new_page = pdf_writer.add_blank_page(width=pdf_width, height=pdf_height)
                new_page.merge_page(original_page)
                logger.debug(f"Merged original content for page {page_index_1_based}")
            except Exception as page_setup_err:
                 logger.error(f"Error setting up or merging original page {page_index_1_based}: {page_setup_err}", exc_info=True)
                 try:
                     pdf_writer.add_blank_page(width=612, height=792)
                     logger.warning(f"Added default blank page for page index {page_index_1_based} due to error.")
                 except Exception as add_blank_fallback_err:
                     logger.error(f"Failed even to add a default blank page for page {page_index_1_based}: {add_blank_fallback_err}")
                 continue

            if page_index_1_based in page_coords:
                packet = io.BytesIO()
                try:
                    c = canvas.Canvas(packet, pagesize=(pdf_width, pdf_height))
                    c.setStrokeColorRGB(1, 0, 0); c.setLineWidth(0.5)
                    for coords in page_coords[page_index_1_based]:
                        try:
                            p = c.beginPath()
                            pdf_height_points = pdf_height
                            p.moveTo(coords[0][0] * inch, pdf_height_points - coords[0][1] * inch)
                            p.lineTo(coords[1][0] * inch, pdf_height_points - coords[1][1] * inch)
                            p.lineTo(coords[2][0] * inch, pdf_height_points - coords[2][1] * inch)
                            p.lineTo(coords[3][0] * inch, pdf_height_points - coords[3][1] * inch)
                            p.close()
                            c.drawPath(p, stroke=1, fill=0)
                        except Exception as draw_err:
                            logger.warning(f"Error drawing path page {page_index_1_based} coords {coords}. Skip highlight. Err: {draw_err}")
                            continue
                    c.save()
                    packet.seek(0)
                    try:
                        overlay_reader = PdfReader(packet)
                        if overlay_reader.pages:
                             overlay_page = overlay_reader.pages[0]
                             new_page.merge_page(overlay_page)
                             logger.debug(f"Merged overlay onto new page {page_index_1_based}")
                        else:
                             logger.warning(f"Highlight overlay page {page_index_1_based} empty/invalid. Skip merge.")
                    except Exception as merge_err:
                        logger.error(f"Error merging highlight overlay onto new page {page_index_1_based}: {merge_err}", exc_info=True)
                finally:
                     packet.close()

        pdf_writer.write(output_pdf)
        output_pdf.seek(0)
        logger.info(f"PDF highlighting completed successfully using pypdf (add_blank strategy) into memory stream.")
        return output_pdf

    except (RuntimeError) as e:
        logger.error(f"Highlighting failed for PDF stream: {e}", exc_info=False)
        try:
            logger.warning("Highlighting failed. Returning original PDF content.")
            original_pdf_stream.seek(0)
            return io.BytesIO(original_pdf_stream.read())
        except Exception as fallback_err:
             logger.error(f"Failed to return original PDF after highlighting error: {fallback_err}")
             raise RuntimeError(f"Highlighting failed, could not retrieve original: {e}") from e
    except Exception as e:
        logger.error(f"Unexpected error during highlighting PDF stream (add_blank strategy): {e}", exc_info=True)
        raise RuntimeError(f"Unexpected error during PDF highlighting: {e}")

# -------------------- Flask Routes --------------------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
@limiter.limit("5/minute")
def process_files():
    start_time = time.time()
    request_ip = get_remote_address()
    logger.info(f"Received /process request from {request_ip}")

    if not azure_form_recognizer_client: abort(503, description="Server error: Form Recognizer unavailable.")

    pdf_file = request.files.get("pdf_file")
    excel_file = request.files.get("excel_file")

    if not pdf_file or not excel_file or not pdf_file.filename or not excel_file.filename:
        abort(400, description="Missing required PDF or Excel file, or filename is empty.")

    safe_pdf_filename = secure_filename(pdf_file.filename)
    safe_excel_filename = secure_filename(excel_file.filename)
    pdf_base, _ = os.path.splitext(safe_pdf_filename)
    run_id = str(uuid.uuid4())

    allowed_pdf_mimetypes = ["application/pdf"]
    allowed_excel_mimetypes = ["application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
    allowed_excel_extensions = [".xls", ".xlsx"]
    pdf_mimetype = pdf_file.mimetype
    excel_mimetype = excel_file.mimetype
    _, pdf_ext = os.path.splitext(safe_pdf_filename)
    _, excel_ext = os.path.splitext(safe_excel_filename)
    if pdf_mimetype not in allowed_pdf_mimetypes: abort(400, description="Invalid PDF file type.")
    if excel_mimetype not in allowed_excel_mimetypes and excel_ext.lower() not in allowed_excel_extensions:
        abort(400, description="Invalid Excel file type (.xls or .xlsx).")

    logger.info(f"Processing run {run_id}: PDF='{safe_pdf_filename}', Excel='{safe_excel_filename}' from {request_ip}")

    pdf_stream_in_memory = None
    excel_stream_in_memory = None
    excel_report_stream = None
    highlighted_pdf_stream = None
    output_zip_stream = None

    try:
        # --- 1. Read input files into memory streams ---
        pdf_file.seek(0)
        pdf_stream_in_memory = io.BytesIO(pdf_file.read())
        excel_file.seek(0)
        excel_stream_in_memory = io.BytesIO(excel_file.read())
        logger.info(f"Run {run_id}: Input files read into memory.")

        # --- 2. Read Terms (from memory stream) ---
        try:
            _, search_terms_lower_set, lower_to_original_map = read_search_terms(excel_stream_in_memory)
        except (ValueError, RuntimeError) as e:
            logger.warning(f"Run {run_id}: Failed read search terms from stream: {e}")
            status_code = 400 if isinstance(e, ValueError) else 500
            abort(status_code, description=str(e))

        # --- 3. Process PDF (FR) (from memory stream) ---
        try:
            results = process_pdf_azure(pdf_stream_in_memory, search_terms_lower_set, lower_to_original_map)
        except (ConnectionError, RuntimeError) as e:
            logger.error(f"Run {run_id}: Azure PDF processing failed from stream: {e}", exc_info=False)
            status_code = 503 if isinstance(e, ConnectionError) else 500
            abort(status_code, description=f"Failed to analyze PDF: {e}")
        except Exception as e:
             logger.error(f"Run {run_id}: Unexpected PDF processing error from stream: {e}", exc_info=True)
             abort(500, description=f"Unexpected PDF analysis error: {e}")

        # --- 4. Generate Excel Report (into memory stream) ---
        try:
            excel_report_stream = generate_excel_report(results)
        except (RuntimeError, Exception) as e:
             logger.error(f"Run {run_id}: Excel generation failed: {e}", exc_info=True)
             abort(500, description=f"Failed to generate analysis report: {e}")

        # --- 5. Highlight PDF (using memory stream, output to memory stream) ---
        try:
            highlighted_pdf_stream = highlight_pdf(pdf_stream_in_memory, results)
        except (RuntimeError, Exception) as e:
             logger.error(f"Run {run_id}: PDF highlighting failed: {e}", exc_info=True)
             abort(500, description=f"Failed to highlight PDF: {e}")

        # --- 6. Create ZIP (from memory streams) ---
        output_zip_stream = io.BytesIO()
        try:
            excel_report_stream.seek(0)
            highlighted_pdf_stream.seek(0)
            with zipfile.ZipFile(output_zip_stream, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"{pdf_base}_analysis.xlsx", excel_report_stream.getvalue())
                zf.writestr(f"{pdf_base}_highlighted.pdf", highlighted_pdf_stream.getvalue())
            output_zip_stream.seek(0)
            logger.info(f"Run {run_id}: Successfully created ZIP in memory.")
        except (zipfile.BadZipFile, Exception) as zip_err:
             logger.error(f"Run {run_id}: Failed create ZIP in memory: {zip_err}", exc_info=True)
             abort(500, description="Failed to create final results ZIP file.")

        # --- 7. Send ZIP (from memory stream) ---
        zip_download_filename = f"{pdf_base}_results.zip"
        logger.info(f"Run {run_id}: Success '{safe_pdf_filename}'. Sending ZIP. Total Time: {time.time() - start_time:.2f}s IP: {request_ip}.")
        return send_file(
            output_zip_stream,
            mimetype="application/zip",
            as_attachment=True,
            download_name=zip_download_filename
        )

    except Exception as e:
        logger.exception(f"Run {run_id}: Unhandled exception /process route IP: {request_ip}:")
        if not request.environ.get('werkzeug.response_started'):
            return jsonify({"error": f"Unexpected server error (Ref: {run_id}). Contact support."}), 500
        else:
             logger.error(f"Run {run_id}: Exception occurred after response started.")
             return None

    finally:
        # Clean up in-memory streams
        if pdf_stream_in_memory: pdf_stream_in_memory.close()
        if excel_stream_in_memory: excel_stream_in_memory.close()
        if excel_report_stream: excel_report_stream.close()
        if highlighted_pdf_stream: highlighted_pdf_stream.close()
        # DO NOT CLOSE output_zip_stream here - Flask/Waitress handles it

# -------------------- Error Handlers --------------------
@app.errorhandler(400)
def bad_request(error):
    description = getattr(error, 'description', "Invalid request.")
    logger.warning(f"Bad Request (400) from {get_remote_address()}: {description}")
    return jsonify({"error": description}), 400

@app.errorhandler(413)
def request_entity_too_large(error):
    max_size_mb = app.config["MAX_CONTENT_LENGTH"] / (1024 * 1024)
    description = getattr(error, 'description', f"File size exceeds limit of {max_size_mb:.1f} MB.")
    logger.warning(f"Request Entity Too Large (413) from {get_remote_address()}. Limit: {max_size_mb:.1f} MB")
    return jsonify(error=description), 413

@app.errorhandler(429)
def ratelimit_handler(error):
    limit_info = getattr(error, 'description', "Rate limit exceeded")
    logger.warning(f"Rate Limit Exceeded (429) for {get_remote_address()}. Limit: {limit_info}")
    return jsonify(error=f"Rate limit exceeded ({limit_info}). Try again later."), 429

@app.errorhandler(500)
def internal_server_error(error):
    run_id_str = ""
    description = getattr(error, 'description', "An internal error occurred.")
    if isinstance(description, str):
        match = re.search(r'([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})', description)
        if match: run_id_str = f" (Ref: {match.group(1)})"
        elif "Run ID:" in description:
             try: run_id_str = f" (Ref: {description.split('Run ID:')[1].split(')')[0].strip()})"
             except IndexError: pass

    original_exception = getattr(error, 'original_exception', error)
    logger.error(f"Internal Server Error (500) from {get_remote_address()}{run_id_str}: {description}", exc_info=original_exception)
    user_message = f"An unexpected internal server error occurred{run_id_str}. Please contact support."
    return jsonify(error=user_message), 500

@app.errorhandler(503)
def service_unavailable(error):
     description = getattr(error, 'description', "Service temporarily unavailable.")
     logger.error(f"Service Unavailable (503) from {get_remote_address()}: {description}")
     return jsonify(error=description), 503

# --- Main Execution (Reads PORT environment variable) ---
if __name__ == "__main__":
    static_folder_path = os.path.join(os.path.dirname(__file__), 'static')
    if not os.path.exists(static_folder_path):
        try: os.makedirs(static_folder_path); logger.info(f"Created 'static' folder.")
        except OSError as e: logger.error(f"Failed to create 'static' folder: {e}")

    # Check Azure Clients Status
    clients_ok = True
    if azure_form_recognizer_client is None: print("\n!! ERROR: Azure Form Recognizer client failed !!"); clients_ok = False
    if azure_blob_service_client is None: print("\n!! WARNING: Azure Storage client failed initialization (may not impact core processing) !!")

    if not clients_ok:
        print("\n!! Critical Azure Form Recognizer client failed. App cannot run. Check logs. !!\n")
        import sys; sys.exit(1)
    else:
         print("\n--- Azure Clients Initialized ---")
         print(f"Form Recognizer: {AZURE_FORM_RECOGNIZER_ENDPOINT}")
         if azure_blob_service_client:
             print(f"Storage Account: {azure_blob_service_client.account_name} / Container: {AZURE_STORAGE_CONTAINER_NAME} (Initialized, but not used in critical path)")
         else:
             print("Storage Account: Client failed initialization.")
         print("---------------------------------\n")

    # --- Get Port from Environment Variable for Azure Web Apps ---
    # Use a default (e.g., 5000) for local running if PORT is not set
    port = int(os.environ.get("PORT", 5000))
    # -------------------------------------------------------------

    print(f"Starting server using Waitress on http://0.0.0.0:{port}") # Use the determined port
    print("Ensure 'epiqlogo.png' is in the 'static' directory.")
    print("\n*** WARNING: Using hardcoded Azure credentials. DO NOT deploy sensitive data this way in production. ***\n")
    print("*** Optimization Applied: Intermediate files processed in memory. ***\n")

    from waitress import serve
    # Use the port variable in the serve call
    serve(app, host='0.0.0.0', port=port, threads=8, channel_timeout=900)
