# -*- coding: utf-8 -*-
"""
Flask Web Application for PDF Keyword Analysis and Highlighting.

This application allows users to upload a PDF file and an Excel file containing
search terms. It uses Azure Form Recognizer to analyze the PDF, identifies occurrences
of the search terms, generates an Excel report with the locations (bounding boxes)
of found keywords, and creates a new PDF with the keywords highlighted. The results
are bundled into a ZIP file and returned to the user.

Security Note: Azure credentials (Endpoint, Key, Storage Connection String) MUST be
configured via environment variables. DO NOT hardcode them in the source code.
"""

# --- Standard Library Imports ---
import os
import io
import zipfile
import logging
import re
import time
import uuid  # For generating unique identifiers for requests/blobs
import traceback # Kept for potential deep debugging if needed

# --- Third-Party Imports ---
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify, abort
from werkzeug.utils import secure_filename
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

# --- Azure SDK Imports ---
from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import HttpResponseError, ServiceRequestError, ResourceNotFoundError
from azure.ai.formrecognizer import FormRecognizerClient
# Keep BlobServiceClient for potential future use or if other parts need it.
# The current primary workflow processes files entirely in memory.
from azure.storage.blob import BlobServiceClient #, BlobClient, ContainerClient (Uncomment if needed)

# --- PDF Processing Imports ---
from pypdf import PdfReader, PdfWriter # Using pypdf for PDF manipulation
from reportlab.pdfgen import canvas   # For creating PDF overlays (highlights)
from reportlab.lib.units import inch  # For unit conversion in PDF generation

# -------------------- Flask App Setup --------------------
app = Flask(__name__, static_folder='static')
# Set maximum upload size (e.g., 500 MB). Adjust as needed.
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

# Rate Limiting: Protects the service from abuse.
# Limits apply per IP address. Adjust limits based on expected usage.
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["200 per day", "50 per hour"] # Example limits
)

# -------------------- Logging Setup ----------------------
# Configure logging format and level.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(name)s - Thread: %(thread)d - %(message)s'
)
# Suppress verbose logs from Azure SDKs unless there's a warning/error.
logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
logging.getLogger("azure.storage.blob").setLevel(logging.WARNING)
# Suppress info logs from pypdf, only show warnings/errors.
logging.getLogger("pypdf").setLevel(logging.WARNING)
# Get logger for this specific application module.
logger = logging.getLogger(__name__)

# -------------------- Configuration from Environment Variables --------------------
# --- Azure Form Recognizer Setup ---
# !! IMPORTANT !! Obtain credentials from environment variables for security.
AZURE_FORM_RECOGNIZER_ENDPOINT = os.environ.get("AZURE_FORM_RECOGNIZER_ENDPOINT")
AZURE_FORM_RECOGNIZER_KEY = os.environ.get("AZURE_FORM_RECOGNIZER_KEY")

azure_form_recognizer_client = None
if not AZURE_FORM_RECOGNIZER_ENDPOINT or not AZURE_FORM_RECOGNIZER_KEY:
    logger.error("FATAL: Azure Form Recognizer Endpoint or Key not found in environment variables.")
    # Application cannot function without Form Recognizer, so we might exit later
else:
    try:
        azure_form_recognizer_client = FormRecognizerClient(
            endpoint=AZURE_FORM_RECOGNIZER_ENDPOINT,
            credential=AzureKeyCredential(AZURE_FORM_RECOGNIZER_KEY),
            logging_enable=False # Disable verbose SDK logging here, use global config
        )
        logger.info("Azure Form Recognizer client initialized successfully.")
    except Exception as e:
        logger.error(f"FATAL: Failed to initialize Azure Form Recognizer client: {e}", exc_info=True)
        azure_form_recognizer_client = None # Ensure it's None if initialization failed

# --- Azure Storage Setup ---
# !! IMPORTANT !! Obtain connection string from environment variables.
# Note: This client initialization is kept, but the primary processing
# workflow (`process_files`) now operates on in-memory streams, not Azure blobs
# for intermediate steps. This client might be used for future features like
# archival or handling very large files if memory constraints become an issue.
AZURE_STORAGE_CONNECTION_STRING = os.environ.get("AZURE_STORAGE_CONNECTION_STRING")
AZURE_STORAGE_CONTAINER_NAME = os.environ.get("AZURE_STORAGE_CONTAINER_NAME", "tempfiles") # Default container name if not set

azure_blob_service_client = None
if not AZURE_STORAGE_CONNECTION_STRING:
    logger.warning("Azure Storage Connection String not found in environment variables. Blob storage features unavailable.")
else:
    try:
        azure_blob_service_client = BlobServiceClient.from_connection_string(
            AZURE_STORAGE_CONNECTION_STRING,
            logging_enable=False # Disable verbose SDK logging here
        )
        logger.info("Azure Blob Storage client initialized.")
        # Optional: Check container access if blob storage is actively used.
        # try:
        #     container_client = azure_blob_service_client.get_container_client(AZURE_STORAGE_CONTAINER_NAME)
        #     container_client.get_container_properties() # Raises exception if container doesn't exist or access denied
        #     logger.info(f"Successfully accessed storage container '{AZURE_STORAGE_CONTAINER_NAME}'.")
        # except Exception as container_ex:
        #     logger.warning(f"Could not access storage container '{AZURE_STORAGE_CONTAINER_NAME}': {container_ex}")
    except Exception as e:
        logger.error(f"Failed to initialize Azure Storage client: {e}", exc_info=True)
        azure_blob_service_client = None # Mark as unavailable

# -------------------- Utility Functions --------------------

def read_search_terms(excel_stream: io.BytesIO):
    """
    Reads search terms from the first column of an Excel file stream.

    Performs case-insensitive matching by converting terms to lowercase
    while retaining the original casing for reporting.

    :param excel_stream: A BytesIO stream containing the Excel file content.
    :return: A tuple containing:
             - list[str]: Original unique search terms.
             - set[str]: Lowercase unique search terms for matching.
             - dict[str, str]: Mapping from lowercase term to original term.
    :raises ValueError: If the Excel file is empty, cannot be parsed, or contains no terms.
    :raises RuntimeError: For unexpected errors during processing.
    """
    logger.info("Reading search terms from provided Excel stream.")
    try:
        excel_stream.seek(0) # Ensure stream pointer is at the beginning
        # Read only the first column, assume no header
        df = pd.read_excel(excel_stream, usecols=[0], header=None, engine='openpyxl')
        # Drop empty rows, convert to string, get unique values, convert to list
        search_terms_original = df.iloc[:, 0].dropna().astype(str).unique().tolist()

        if not search_terms_original:
            raise ValueError("No search terms found in the first column of the Excel file.")

        search_terms_lower_set = set()
        lower_to_original_map = {}
        for term in search_terms_original:
            lower_term = term.lower()
            search_terms_lower_set.add(lower_term)
            # Store the first encountered original casing for each lower case term
            if lower_term not in lower_to_original_map:
                 lower_to_original_map[lower_term] = term

        logger.info(f"Successfully read {len(search_terms_original)} unique search terms.")
        return search_terms_original, search_terms_lower_set, lower_to_original_map

    except ValueError as ve: # Catch specific pandas/value errors
         logger.error(f"Error reading or processing Excel stream: {ve}", exc_info=False) # No need for stack trace here
         raise ValueError(f"Invalid Excel file content or format: {ve}")
    except Exception as e:
        logger.error(f"Unexpected error reading Excel stream: {e}", exc_info=True)
        raise RuntimeError(f"An unexpected error occurred while reading the Excel file: {e}")

def generate_excel_report(results: list[dict]) -> io.BytesIO:
    """
    Generates an Excel report from the analysis results into a BytesIO stream.

    The report includes the found keyword, page number, and bounding box coordinates.

    :param results: A list of dictionaries, where each dict represents a found keyword
                    and contains at least 'Keyword', 'Page', and 'WordObject' (with bounding_box).
    :return: A BytesIO stream containing the generated Excel file (.xlsx).
    :raises RuntimeError: If there's an error during Excel generation.
    """
    logger.info(f"Generating Excel report for {len(results)} found items.")
    report_columns = ["Keyword", "Page", "X1", "Y1", "X2", "Y2", "X3", "Y3", "X4", "Y4"]

    if not results:
        # Create empty DataFrame with correct columns if no results
        df = pd.DataFrame(columns=report_columns)
        logger.warning("No results found, generating an empty Excel report.")
    else:
        formatted_results = []
        for item in results:
            # Validate structure before accessing deeply nested attributes
            if not isinstance(item, dict) or 'WordObject' not in item or 'Keyword' not in item or 'Page' not in item:
                logger.warning(f"Skipping invalid result item structure: {item}")
                continue

            word_obj = item.get('WordObject')
            if not hasattr(word_obj, 'bounding_box'):
                 logger.warning(f"Skipping result item for keyword '{item.get('Keyword', 'N/A')}' on page {item.get('Page', 'N/A')} due to missing bounding_box.")
                 continue

            coords = word_obj.bounding_box
            # Azure Form Recognizer typically returns 4 points (8 coordinates)
            if coords and len(coords) == 4:
                # Ensure all points have x and y attributes
                if all(hasattr(p, 'x') and hasattr(p, 'y') for p in coords):
                    formatted_results.append({
                        "Keyword": item['Keyword'],
                        "Page": item['Page'],
                        "X1": coords[0].x, "Y1": coords[0].y,
                        "X2": coords[1].x, "Y2": coords[1].y,
                        "X3": coords[2].x, "Y3": coords[2].y,
                        "X4": coords[3].x, "Y4": coords[3].y,
                    })
                else:
                    logger.warning(f"Keyword '{item['Keyword']}' on page {item['Page']} has malformed coordinate points. Skipping.")
            else:
                 logger.warning(f"Keyword '{item['Keyword']}' on page {item['Page']} has unexpected bounding box length ({len(coords)}). Skipping.")

        if not formatted_results:
             df = pd.DataFrame(columns=report_columns)
             logger.warning("No valid results could be formatted, generating an empty Excel report.")
        else:
             df = pd.DataFrame(formatted_results)

    output_stream = io.BytesIO()
    try:
        # Use openpyxl engine for .xlsx format
        with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="AnalysisResults")
        output_stream.seek(0) # Rewind the stream to the beginning for reading
        logger.info("Excel report generated successfully into memory stream.")
        return output_stream
    except Exception as e:
        logger.error(f"Error generating Excel report: {e}", exc_info=True)
        raise RuntimeError(f"Failed to generate the Excel analysis report: {e}")

def process_pdf_azure(pdf_stream: io.BytesIO, search_terms_lower_set: set[str], lower_to_original_map: dict[str, str], polling_interval_seconds: int = 5) -> list[dict]:
    """
    Analyzes a PDF stream using Azure Form Recognizer to find occurrences of search terms.

    Uses the 'recognize_content' (Layout) model which extracts text, lines, words,
    and their bounding boxes. Performs case-insensitive matching.

    :param pdf_stream: A BytesIO stream containing the PDF content.
    :param search_terms_lower_set: A set of lowercase search terms.
    :param lower_to_original_map: A dictionary mapping lowercase terms to their original casing.
    :param polling_interval_seconds: How often to check the status of the Azure job.
    :return: A list of dictionaries, each representing a found keyword instance,
             containing 'Keyword', 'Page', and the 'WordObject' from Form Recognizer.
    :raises ConnectionError: If the Form Recognizer client is not available or if there's
                             a network/authentication issue with Azure.
    :raises RuntimeError: For unexpected errors during processing or if the Azure job fails.
    """
    if not azure_form_recognizer_client:
        raise ConnectionError("Azure Form Recognizer client is not initialized or unavailable.")

    logger.info("Starting Azure Form Recognizer analysis on the provided PDF stream.")
    pdf_stream.seek(0) # Ensure stream is at the beginning for Azure SDK

    try:
        # Start the recognition process. 'recognize_content' extracts layout and text.
        # content_type is crucial for the service to interpret the stream correctly.
        poller = azure_form_recognizer_client.begin_recognize_content(
            pdf_stream,
            content_type='application/pdf',
            logging_enable=False # Use global logging config
        )
        logger.info(f"Azure analysis job submitted. Polling for results every {polling_interval_seconds} seconds...")

        start_poll_time = time.time()
        # Wait for the asynchronous operation to complete
        # poller.result() will block until completion, but adding a loop provides
        # more control and logging opportunities, especially for long documents.
        while not poller.done():
            try:
                status = poller.status()
                logger.debug(f"Polling Azure... Current status: {status}")
                if poller.done():
                    break # Exit loop if already done
                poller.wait(timeout=polling_interval_seconds) # Wait for specified interval
            except TimeoutError:
                # This can happen if the timeout expires; just continue polling
                logger.debug(f"Polling timeout after {polling_interval_seconds}s, checking status again...")
                continue
            except (HttpResponseError, ServiceRequestError) as poll_err:
                # Errors during the polling process itself (network, temporary service issues)
                logger.error(f"Error while polling Azure job status: {poll_err}", exc_info=True)
                raise ConnectionError(f"Communication failed while waiting for Azure analysis: {poll_err}") from poll_err
            except Exception as poll_err:
                 # Catch unexpected errors during the wait/status check
                 logger.error(f"Unexpected error during Azure polling wait: {poll_err}", exc_info=True)
                 raise RuntimeError(f"Unexpected error while waiting for Azure analysis: {poll_err}") from poll_err

        elapsed_time = time.time() - start_poll_time
        logger.info(f"Azure polling completed in {elapsed_time:.2f} seconds.")

        final_status = poller.status()
        logger.info(f"Azure Form Recognizer job finished with status: {final_status}")

        # Check the final status and get the result
        if final_status.lower() == "succeeded":
            # Get the result pages (FormPage instances)
            result = poller.result()
            if not result:
                 logger.warning("Azure analysis succeeded but returned no result pages.")
                 return []
        else:
            # If the job failed, attempt to get the result to potentially raise a more specific error
            error_message = f"Azure Form Recognizer job failed with status '{final_status}'."
            try:
                poller.result() # This should raise an exception if the job failed
            except HttpResponseError as http_err:
                # Log details if available from the HTTP error
                logger.error(f"Azure job failed. Status Code: {http_err.status_code}, Reason: {http_err.reason}, Details: {http_err.message}")
                error_message = f"Azure analysis failed (Status Code: {http_err.status_code}). Check Azure portal for details."
            except Exception as e:
                logger.error(f"Azure job failed, and error retrieving result details: {e}", exc_info=True)
                error_message = f"Azure analysis failed with status '{final_status}'. Unable to retrieve detailed error."
            raise RuntimeError(error_message)

        # Process the results: Iterate through pages, lines, and words
        pdf_results = []
        # Regex to remove common non-alphanumeric characters (excluding whitespace) for cleaner matching
        # Adjust this regex based on the specific types of characters to ignore in keywords
        non_word_char_regex = re.compile(r'[^\w\s]')

        for page in result:
            if not page.lines: continue # Skip pages with no extracted lines

            for line in page.lines:
                if not line.words: continue # Skip lines with no extracted words

                for word in line.words:
                    # Clean the extracted word: remove punctuation, convert to lowercase
                    # This allows for more robust matching against the search terms
                    cleaned_word_text = non_word_char_regex.sub('', word.text).lower()

                    # Check if the cleaned word matches any of the lowercase search terms
                    if cleaned_word_text in search_terms_lower_set:
                        # Retrieve the original casing of the search term
                        original_term = lower_to_original_map.get(cleaned_word_text, word.text) # Fallback to word's text if somehow not in map
                        # Store the finding along with context
                        pdf_results.append({
                            "Keyword": original_term,
                            "Page": page.page_number, # 1-based page number
                            "WordObject": word # Store the full word object for coordinates etc.
                        })

        logger.info(f"Found {len(pdf_results)} keyword instances in the PDF using Azure Form Recognizer.")
        return pdf_results

    except (ConnectionError, RuntimeError) as e:
        # Re-raise known error types for specific handling upstream
        raise e
    except Exception as e:
        # Catch any other unexpected exceptions during the process
        logger.error(f"Unexpected error during Azure PDF processing: {e}", exc_info=True)
        raise RuntimeError(f"An unexpected error occurred during PDF analysis with Azure: {e}") from e


def highlight_pdf(original_pdf_stream: io.BytesIO, results_data: list[dict]) -> io.BytesIO:
    """
    Highlights keywords in a PDF using bounding box data from analysis results.

    Creates a new PDF with highlights overlaid on the original content using
    pypdf and reportlab. Operates entirely on BytesIO streams.

    Strategy:
    1. Reads the original PDF page by page (using pypdf).
    2. For each page, creates a blank page of the same dimensions in a new PDF (using pypdf).
    3. Merges the original page content onto the new blank page.
    4. If highlights are needed for this page:
       - Creates a temporary PDF overlay containing only the highlight rectangles (using reportlab).
       - Merges this overlay onto the new page containing the original content.
    5. Writes the final combined PDF to a BytesIO stream.

    :param original_pdf_stream: A BytesIO stream containing the original PDF content.
    :param results_data: A list of dictionaries from `process_pdf_azure`, containing
                         keyword locations ('Page', 'WordObject' with 'bounding_box').
    :return: A BytesIO stream containing the highlighted PDF content. If highlighting fails,
             it attempts to return the original PDF content.
    :raises RuntimeError: For unexpected errors during PDF manipulation or highlighting.
    """
    if not results_data:
        logger.warning("No results data provided for highlighting. Returning original PDF content.")
        original_pdf_stream.seek(0)
        # Return a *copy* of the original stream content to avoid issues with stream closing
        return io.BytesIO(original_pdf_stream.read())

    logger.info(f"Starting PDF highlighting process for {len(results_data)} found items.")

    # --- Organize coordinates by page number ---
    # Page numbers from Form Recognizer are 1-based.
    page_coords = {}
    for item in results_data:
        try:
            page_num_1_based = item['Page']
            word_obj = item['WordObject']
            # Bounding box from Form Recognizer is usually a list of 4 Point objects (x, y)
            coords_raw = word_obj.bounding_box
            keyword_text = item.get('Keyword', 'N/A') # For logging

            if coords_raw and len(coords_raw) == 4:
                # Validate and convert coordinates to float tuples (x, y)
                coords = [(float(p.x), float(p.y)) for p in coords_raw]
                if page_num_1_based not in page_coords:
                    page_coords[page_num_1_based] = []
                page_coords[page_num_1_based].append(coords)
            else:
                logger.warning(f"Invalid or missing bounding box for keyword '{keyword_text}' on page {page_num_1_based}. Skipping highlight for this instance.")
                continue
        except (AttributeError, TypeError, ValueError, KeyError) as ex:
            logger.warning(f"Error processing coordinates for item on page {item.get('Page', 'N/A')}. Keyword: '{item.get('Keyword', 'N/A')}'. Error: {ex}. Skipping highlight for this instance.")
            continue

    if not page_coords:
        logger.warning("No valid coordinates found in results data. Returning original PDF content.")
        original_pdf_stream.seek(0)
        return io.BytesIO(original_pdf_stream.read())


    output_pdf_stream = io.BytesIO()
    try:
        original_pdf_stream.seek(0) # Ensure original stream is at the start
        original_pdf_reader = PdfReader(original_pdf_stream)
        pdf_writer = PdfWriter() # Creates a new PDF document in memory
        num_pages = len(original_pdf_reader.pages)

        logger.info(f"Processing {num_pages} pages from PDF stream for highlighting.")

        # Iterate through pages of the original PDF (0-based index)
        for page_index in range(num_pages):
            page_num_1_based = page_index + 1
            original_page = original_pdf_reader.pages[page_index]

            # --- Create a new page and merge original content ---
            try:
                # Get original page dimensions (MediaBox is standard)
                page_box = original_page.mediabox
                pdf_width_pt = float(page_box.width)  # Width in points
                pdf_height_pt = float(page_box.height) # Height in points

                # Add a blank page with the same dimensions to the writer
                # Note: pypdf's add_blank_page uses points directly.
                new_page = pdf_writer.add_blank_page(width=pdf_width_pt, height=pdf_height_pt)

                # Merge the original page content onto the newly created blank page
                new_page.merge_page(original_page)
                logger.debug(f"Merged original content for page {page_num_1_based}")

            except Exception as page_setup_err:
                 logger.error(f"Error setting up or merging original page {page_num_1_based}: {page_setup_err}", exc_info=True)
                 # As a fallback, add a standard US Letter size page if setup fails. Content will be lost.
                 try:
                     pdf_writer.add_page(original_page) # Try adding the original page directly if merge failed
                     logger.warning(f"Fallback: Added original page {page_num_1_based} directly due to merge error (highlights might be lost/misaligned).")
                 except Exception as add_page_fallback_err:
                      logger.error(f"Critical: Failed even to add original page {page_num_1_based} after setup error: {add_page_fallback_err}. Skipping page.")
                 continue # Skip highlighting for this problematic page

            # --- Create and merge highlight overlay if coordinates exist for this page ---
            if page_num_1_based in page_coords:
                overlay_stream = io.BytesIO()
                try:
                    # Create an overlay canvas using reportlab
                    # Important: ReportLab uses points (1/72 inch) as default unit.
                    # Form Recognizer coordinates are typically in inches relative to top-left.
                    # We need to convert FR coordinates (inches, top-left origin) to
                    # ReportLab coordinates (points, bottom-left origin).
                    c = canvas.Canvas(overlay_stream, pagesize=(pdf_width_pt, pdf_height_pt))
                    c.setStrokeColorRGB(1, 0, 0)  # Red color for rectangle border
                    c.setFillColorRGB(1, 1, 0, alpha=0.3) # Semi-transparent yellow fill (optional)
                    c.setLineWidth(1) # Line width in points

                    for coords_list in page_coords[page_num_1_based]:
                        try:
                            # Assuming coords_list is a list of 4 (x, y) tuples in inches from FR
                            # Convert inches to points (multiply by 72)
                            # Adjust Y coordinate: PDF Y = PageHeight - FR_Y
                            # Coords order: top-left, top-right, bottom-right, bottom-left (usually)
                            tl_x, tl_y = coords_list[0][0] * inch, pdf_height_pt - (coords_list[0][1] * inch)
                            tr_x, tr_y = coords_list[1][0] * inch, pdf_height_pt - (coords_list[1][1] * inch)
                            br_x, br_y = coords_list[2][0] * inch, pdf_height_pt - (coords_list[2][1] * inch)
                            bl_x, bl_y = coords_list[3][0] * inch, pdf_height_pt - (coords_list[3][1] * inch)

                            # Draw a polygon using the converted points
                            # ReportLab uses bottom-left origin.
                            path = c.beginPath()
                            path.moveTo(tl_x, tl_y)
                            path.lineTo(tr_x, tr_y)
                            path.lineTo(br_x, br_y)
                            path.lineTo(bl_x, bl_y)
                            path.close()
                            # Set fill=1 for filled rectangle, stroke=1 for border
                            c.drawPath(path, stroke=1, fill=1)

                        except IndexError:
                            logger.warning(f"Skipping a highlight on page {page_num_1_based} due to incomplete coordinates: {coords_list}")
                        except Exception as draw_err:
                            logger.warning(f"Error drawing highlight path on page {page_num_1_based} for coords {coords_list}. Error: {draw_err}. Skipping this highlight.")
                            continue # Skip this specific highlight, continue with others

                    c.save() # Finalize the overlay PDF page
                    overlay_stream.seek(0)

                    # Merge the overlay onto the new page
                    try:
                        overlay_pdf_reader = PdfReader(overlay_stream)
                        if overlay_pdf_reader.pages:
                             # Assuming single-page overlay
                            overlay_page = overlay_pdf_reader.pages[0]
                            new_page.merge_page(overlay_page)
                            logger.debug(f"Merged highlight overlay onto page {page_num_1_based}")
                        else:
                             logger.warning(f"Generated highlight overlay for page {page_num_1_based} was empty or invalid. Skipping merge.")
                    except Exception as merge_err:
                        logger.error(f"Error merging highlight overlay onto page {page_num_1_based}: {merge_err}", exc_info=True)
                        # Continue to next page, highlights for this page are lost

                finally:
                     overlay_stream.close() # Ensure overlay stream is closed

        # --- Write the final PDF to the output stream ---
        pdf_writer.write(output_pdf_stream)
        output_pdf_stream.seek(0) # Rewind stream for sending
        logger.info("PDF highlighting completed successfully using pypdf and reportlab.")
        return output_pdf_stream

    except Exception as e:
        logger.error(f"Unexpected error during PDF highlighting process: {e}", exc_info=True)
        # Attempt to return the original PDF as a fallback
        try:
            logger.warning("Highlighting failed. Attempting to return original PDF content.")
            original_pdf_stream.seek(0)
            return io.BytesIO(original_pdf_stream.read())
        except Exception as fallback_err:
             logger.error(f"Critical: Failed to return original PDF after highlighting error: {fallback_err}")
             # Raise a runtime error indicating complete failure
             raise RuntimeError(f"PDF highlighting failed, and could not return original PDF content due to: {e}") from e

# -------------------- Flask Routes --------------------

@app.route("/", methods=["GET"])
def index():
    """Renders the main upload page."""
    return render_template("index.html")

@app.route("/process", methods=["POST"])
@limiter.limit("10 per minute") # Adjust rate limit as needed
def process_files():
    """
    Handles the file uploads, processing, and returning the results.

    Workflow:
    1. Validates input files (PDF and Excel).
    2. Reads files into memory streams.
    3. Reads search terms from the Excel stream.
    4. Sends the PDF stream to Azure Form Recognizer for analysis.
    5. Generates an Excel report of the findings (in memory).
    6. Highlights the original PDF based on findings (in memory).
    7. Zips the Excel report and highlighted PDF (in memory).
    8. Sends the ZIP file back to the user.
    """
    start_time = time.time()
    request_ip = get_remote_address() or "Unknown IP"
    # Generate a unique ID for this specific request for tracking/logging
    run_id = str(uuid.uuid4())
    logger.info(f"Processing request Run ID: {run_id} from IP: {request_ip}")

    # --- Essential Service Check ---
    if not azure_form_recognizer_client:
        logger.error(f"Run ID: {run_id}: Aborting request - Azure Form Recognizer client is unavailable.")
        # 503 Service Unavailable is appropriate here
        abort(503, description="Server configuration error: Form Recognizer service is unavailable. Please contact support.")

    # --- File Upload Validation ---
    pdf_file = request.files.get("pdf_file")
    excel_file = request.files.get("excel_file")

    if not pdf_file or not excel_file:
        logger.warning(f"Run ID: {run_id}: Bad Request - Missing PDF or Excel file from {request_ip}.")
        abort(400, description="Missing required files. Please upload both a PDF and an Excel file.")
    if not pdf_file.filename or not excel_file.filename:
         logger.warning(f"Run ID: {run_id}: Bad Request - Uploaded file is missing a filename from {request_ip}.")
         abort(400, description="One or both uploaded files are missing a filename.")

    # Sanitize filenames for security
    safe_pdf_filename = secure_filename(pdf_file.filename)
    safe_excel_filename = secure_filename(excel_file.filename)
    if not safe_pdf_filename or not safe_excel_filename:
        logger.warning(f"Run ID: {run_id}: Bad Request - Filenames potentially insecure or empty after sanitization: PDF='{safe_pdf_filename}', Excel='{safe_excel_filename}' from {request_ip}.")
        abort(400, description="Invalid filenames provided.")

    # Get base name for output files
    pdf_base, pdf_ext = os.path.splitext(safe_pdf_filename)
    _, excel_ext = os.path.splitext(safe_excel_filename)

    # --- MIME Type and Extension Validation ---
    # Define allowed types/extensions
    allowed_pdf_mimetypes = ["application/pdf"]
    allowed_excel_mimetypes = ["application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
    allowed_excel_extensions = [".xls", ".xlsx"]

    # Check PDF type
    pdf_mimetype = pdf_file.mimetype
    if pdf_mimetype not in allowed_pdf_mimetypes or pdf_ext.lower() != ".pdf":
        logger.warning(f"Run ID: {run_id}: Bad Request - Invalid PDF file type ('{pdf_mimetype}', '{pdf_ext}') from {request_ip}.")
        abort(400, description="Invalid file type. Please upload a valid PDF file (.pdf).")

    # Check Excel type (more lenient, checks MIME or extension)
    excel_mimetype = excel_file.mimetype
    if excel_mimetype not in allowed_excel_mimetypes and excel_ext.lower() not in allowed_excel_extensions:
        logger.warning(f"Run ID: {run_id}: Bad Request - Invalid Excel file type ('{excel_mimetype}', '{excel_ext}') from {request_ip}.")
        abort(400, description="Invalid file type. Please upload a valid Excel file (.xls or .xlsx).")

    logger.info(f"Run ID: {run_id}: Processing PDF='{safe_pdf_filename}', Excel='{safe_excel_filename}'")

    # Initialize stream variables (important for the finally block)
    pdf_stream_in_memory = None
    excel_stream_in_memory = None
    excel_report_stream = None
    highlighted_pdf_stream = None
    output_zip_stream = None

    try:
        # --- 1. Read input files into memory streams ---
        # This avoids saving temporary files to disk or blob storage for processing
        pdf_file.seek(0)
        pdf_stream_in_memory = io.BytesIO(pdf_file.read())
        excel_file.seek(0)
        excel_stream_in_memory = io.BytesIO(excel_file.read())
        logger.info(f"Run ID: {run_id}: Input files read into memory.")

        # --- 2. Read Search Terms ---
        try:
            _, search_terms_lower_set, lower_to_original_map = read_search_terms(excel_stream_in_memory)
        except ValueError as e:
            logger.warning(f"Run ID: {run_id}: Failed to read search terms from '{safe_excel_filename}': {e}")
            # 400 Bad Request for invalid Excel content
            abort(400, description=f"Error reading search terms from Excel: {e}")
        except RuntimeError as e:
            logger.error(f"Run ID: {run_id}: Runtime error reading search terms from '{safe_excel_filename}': {e}", exc_info=True)
            # 500 Internal Server Error for unexpected issues
            abort(500, description=f"Server error reading Excel file (Ref: {run_id}).")

        # --- 3. Process PDF with Azure Form Recognizer ---
        try:
            results = process_pdf_azure(pdf_stream_in_memory, search_terms_lower_set, lower_to_original_map)
        except ConnectionError as e:
            logger.error(f"Run ID: {run_id}: Connection error during Azure PDF processing for '{safe_pdf_filename}': {e}", exc_info=False)
            # 503 Service Unavailable if connection to Azure fails
            abort(503, description=f"Failed to connect to analysis service: {e} (Ref: {run_id}).")
        except RuntimeError as e:
            logger.error(f"Run ID: {run_id}: Runtime error during Azure PDF processing for '{safe_pdf_filename}': {e}", exc_info=False) # RuntimeErrors from func are usually descriptive
             # 500 Internal Server Error or potentially a specific Azure failure code if parsed
            abort(500, description=f"PDF analysis failed: {e} (Ref: {run_id}).")
        except Exception as e:
             # Catch any other unexpected errors during PDF processing
             logger.error(f"Run ID: {run_id}: Unexpected error during PDF processing stage for '{safe_pdf_filename}': {e}", exc_info=True)
             abort(500, description=f"Unexpected server error during PDF analysis (Ref: {run_id}).")

        # --- 4. Generate Excel Report ---
        try:
            excel_report_stream = generate_excel_report(results)
        except RuntimeError as e:
             logger.error(f"Run ID: {run_id}: Failed to generate Excel report for '{safe_pdf_filename}': {e}", exc_info=True)
             abort(500, description=f"Failed to generate analysis report (Ref: {run_id}).")
        except Exception as e:
             # Catch any other unexpected errors during Excel generation
             logger.error(f"Run ID: {run_id}: Unexpected error during Excel report generation for '{safe_pdf_filename}': {e}", exc_info=True)
             abort(500, description=f"Unexpected server error generating report (Ref: {run_id}).")

        # --- 5. Highlight PDF ---
        try:
            # Pass the *original* PDF stream again for highlighting
            highlighted_pdf_stream = highlight_pdf(pdf_stream_in_memory, results)
        except RuntimeError as e:
             logger.error(f"Run ID: {run_id}: Failed to highlight PDF '{safe_pdf_filename}': {e}", exc_info=True)
             abort(500, description=f"Failed to generate highlighted PDF (Ref: {run_id}).")
        except Exception as e:
             # Catch any other unexpected errors during highlighting
             logger.error(f"Run ID: {run_id}: Unexpected error during PDF highlighting stage for '{safe_pdf_filename}': {e}", exc_info=True)
             abort(500, description=f"Unexpected server error during PDF highlighting (Ref: {run_id}).")

        # --- 6. Create ZIP file in memory ---
        output_zip_stream = io.BytesIO()
        try:
            # Ensure streams are ready to be read from the beginning
            excel_report_stream.seek(0)
            highlighted_pdf_stream.seek(0)

            # Create ZIP file
            with zipfile.ZipFile(output_zip_stream, "w", zipfile.ZIP_DEFLATED) as zf:
                # Add the Excel report
                zf.writestr(f"{pdf_base}_analysis_results.xlsx", excel_report_stream.getvalue())
                # Add the highlighted PDF
                zf.writestr(f"{pdf_base}_highlighted.pdf", highlighted_pdf_stream.getvalue())

            output_zip_stream.seek(0) # Rewind the ZIP stream for sending
            logger.info(f"Run ID: {run_id}: Successfully created results ZIP file in memory.")

        except (zipfile.BadZipFile, Exception) as zip_err:
             logger.error(f"Run ID: {run_id}: Failed to create ZIP archive in memory for '{safe_pdf_filename}': {zip_err}", exc_info=True)
             abort(500, description=f"Failed to create results ZIP file (Ref: {run_id}).")

        # --- 7. Send ZIP file to user ---
        zip_download_filename = f"{pdf_base}_analysis_results.zip"
        processing_time = time.time() - start_time
        logger.info(f"Run ID: {run_id}: Successfully processed '{safe_pdf_filename}'. Sending results ZIP '{zip_download_filename}'. Total Time: {processing_time:.2f}s. IP: {request_ip}.")

        return send_file(
            output_zip_stream,
            mimetype="application/zip",
            as_attachment=True,
            download_name=zip_download_filename
            # Flask and the WSGI server (e.g., Waitress) handle closing this stream
        )

    except Exception as e:
        # Catch-all for unexpected errors not handled by specific aborts
        # Log the full exception trace
        logger.exception(f"Run ID: {run_id}: Unhandled exception in /process route. IP: {request_ip}")
        # Avoid sending a response if one has already started (e.g., by send_file)
        if not request.environ.get('werkzeug.response_started'):
            # Return a generic 500 error to the client
            return jsonify({"error": f"An unexpected server error occurred. Please try again later or contact support (Ref: {run_id})."}), 500
        else:
             # If response started, we can't send a new one, just log the error.
             logger.error(f"Run ID: {run_id}: Exception occurred after response headers were sent.")
             # Return None or let the WSGI server handle the broken connection
             return None # Explicitly return None

    finally:
        # --- Cleanup: Ensure all opened memory streams are closed ---
        # This is crucial to free up memory, especially under load.
        # The 'output_zip_stream' is handled by Flask/send_file, so DO NOT close it here.
        if pdf_stream_in_memory:
            pdf_stream_in_memory.close()
        if excel_stream_in_memory:
            excel_stream_in_memory.close()
        if excel_report_stream:
            excel_report_stream.close()
        if highlighted_pdf_stream:
            highlighted_pdf_stream.close()
        logger.debug(f"Run ID: {run_id}: Cleaned up in-memory streams.")


# -------------------- Error Handlers --------------------
# These handlers provide consistent JSON error responses for client-side handling.

@app.errorhandler(400) # Bad Request
def bad_request_error(error):
    """Handles client errors like missing files or invalid input format."""
    description = getattr(error, 'description', "Invalid request.")
    # Log the specific error leading to the 400
    logger.warning(f"Bad Request (400) from {get_remote_address() or 'Unknown IP'}: {description}")
    return jsonify(error=description), 400

@app.errorhandler(413) # Payload Too Large
def request_entity_too_large_error(error):
    """Handles errors when uploaded files exceed the configured size limit."""
    max_size_mb = app.config["MAX_CONTENT_LENGTH"] / (1024 * 1024)
    description = f"File size exceeds the server limit of {max_size_mb:.1f} MB."
    logger.warning(f"Request Entity Too Large (413) from {get_remote_address() or 'Unknown IP'}. Limit: {max_size_mb:.1f} MB")
    return jsonify(error=description), 413

@app.errorhandler(429) # Too Many Requests
def ratelimit_error(error):
    """Handles errors when a client exceeds the defined rate limits."""
    limit_info = getattr(error, 'description', "Rate limit exceeded")
    logger.warning(f"Rate Limit Exceeded (429) for {get_remote_address() or 'Unknown IP'}. Limit details: {limit_info}")
    # Provide a user-friendly message
    return jsonify(error=f"You have exceeded the request limit ({limit_info}). Please try again later."), 429

@app.errorhandler(500) # Internal Server Error
def internal_server_error(error):
    """Handles unexpected server errors."""
    # Attempt to extract Run ID if included in the description by abort()
    run_id_str = ""
    description = getattr(error, 'description', "An internal error occurred.")
    if isinstance(description, str):
        # Simple search for UUID pattern or specific "Ref:" marker
        match = re.search(r'([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})', description)
        ref_match = re.search(r'\(Ref: (.*?)\)', description)
        if match:
            run_id_str = f" (Ref: {match.group(1)})"
        elif ref_match:
            run_id_str = f" (Ref: {ref_match.group(1)})"

    # Log the detailed error including stack trace if available
    original_exception = getattr(error, 'original_exception', error)
    logger.error(f"Internal Server Error (500) from {get_remote_address() or 'Unknown IP'}{run_id_str}. Description: {description}", exc_info=original_exception)

    # Return a generic error message to the client, including the reference ID
    user_message = f"An unexpected internal server error occurred{run_id_str}. Please contact support if the issue persists."
    return jsonify(error=user_message), 500

@app.errorhandler(503) # Service Unavailable
def service_unavailable_error(error):
     """Handles errors when dependent services (like Azure) are unavailable."""
     description = getattr(error, 'description', "Service temporarily unavailable.")
     logger.error(f"Service Unavailable (503) from {get_remote_address() or 'Unknown IP'}: {description}")
     # Include reference ID if available in description
     run_id_str = ""
     if isinstance(description, str):
         ref_match = re.search(r'\(Ref: (.*?)\)', description)
         if ref_match: run_id_str = f" (Ref: {ref_match.group(1)})"
     user_message = f"{description}{run_id_str}. Please try again later."
     return jsonify(error=user_message), 503

# -------------------- Main Execution --------------------
if __name__ == "__main__":
    # --- Pre-run Checks ---
    # Ensure static folder exists (Flask usually handles this, but good practice)
    static_folder_path = os.path.join(os.path.dirname(__file__), 'static')
    if not os.path.exists(static_folder_path):
        try:
            os.makedirs(static_folder_path)
            logger.info(f"Created 'static' directory at: {static_folder_path}")
        except OSError as e:
            logger.error(f"Failed to create 'static' directory: {e}")

    # Check critical Azure client status
    clients_ok = True
    if azure_form_recognizer_client is None:
        print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("!! CRITICAL ERROR: Azure Form Recognizer client failed to initialize.       !!")
        print("!! The application cannot function without it.                            !!")
        print("!! Ensure AZURE_FORM_RECOGNIZER_ENDPOINT and AZURE_FORM_RECOGNIZER_KEY    !!")
        print("!! environment variables are set correctly. Check logs for details.       !!")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n")
        clients_ok = False
    else:
        print("\n--- Azure Form Recognizer Client Initialized ---")
        print(f"  Endpoint: {AZURE_FORM_RECOGNIZER_ENDPOINT}")
        print("----------------------------------------------")

    if azure_blob_service_client is None:
        print("\n--- Azure Storage Client Status ---")
        print("!! WARNING: Azure Storage client failed to initialize or is not configured. !!")
        print("!! Blob storage features (if any) will be unavailable.                    !!")
        print("!! Check AZURE_STORAGE_CONNECTION_STRING environment variable if needed.    !!")
        print("-----------------------------------\n")
    else:
        print("\n--- Azure Storage Client Initialized ---")
        print(f"  Account Name: {azure_blob_service_client.account_name}")
        print(f"  Target Container (Default/Env): {AZURE_STORAGE_CONTAINER_NAME}")
        print("  (Note: Client initialized, but core processing uses in-memory streams)")
        print("------------------------------------\n")


    if not clients_ok:
        # Exit if critical components failed
        import sys
        sys.exit(1)

    # --- Server Configuration ---
    # Get port from environment variable (standard for PaaS like Azure Web Apps)
    # Default to 5000 for local development if PORT environment variable is not set.
    port = int(os.environ.get("PORT", 5000))

    # Use Waitress, a production-quality WSGI server.
    # Good choice for cross-platform compatibility and stability compared to Flask's dev server.
    print(f"\n--- Starting Application Server ---")
    print(f" Using Waitress WSGI Server")
    print(f" Listening on: http://0.0.0.0:{port}")
    print(f" Maximum upload size: {app.config['MAX_CONTENT_LENGTH'] / (1024*1024):.1f} MB")
    print(f" Static files served from: '{app.config['static_folder']}'")
    print(f" Ensure required static assets (e.g., logos) are in the static folder.")
    print("-------------------------------------\n")
    print("*** IMPORTANT: Azure credentials configured via environment variables. ***")
    print("*** Application ready. ***\n")

    from waitress import serve
    serve(
        app,
        host='0.0.0.0', # Listen on all available network interfaces
        port=port,
        threads=8,      # Number of worker threads (adjust based on server resources/load)
        channel_timeout=600 # Max time (seconds) to wait for data on a connection (adjust for long uploads/processing)
    )
