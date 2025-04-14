# -*- coding: utf-8 -*-
"""
Flask application for processing PDF lab notebooks.
Takes a PDF and an Excel file (containing keywords) as input.
Uses Azure Form Recognizer to find keyword occurrences in the PDF.
Generates an Excel report listing the locations of keywords and a highlighted PDF.
Returns both files zipped together.
"""

# --- Core and System Libraries ---
import os                   # For interacting with the operating system (e.g., file paths, environment variables)
import io                   # For handling in-memory byte streams (like file uploads/downloads)
import zipfile              # For creating ZIP archives
import logging              # For application logging
import pandas as pd         # For reading Excel files and creating reports
import re                   # For regular expressions (used in text cleaning and error parsing)
import traceback            # For detailed error logging (though less used with custom handlers)
import time                 # For timing operations and polling delays
import uuid                 # For generating unique identifiers (e.g., run IDs, potential future blob names)

# --- Flask Framework and Extensions ---
from flask import Flask, render_template, request, send_file, jsonify, abort # Core Flask components for web server, templates, requests, file sending, JSON responses, and error handling
from werkzeug.utils import secure_filename      # For sanitizing uploaded filenames
from flask_limiter import Limiter               # For rate limiting API requests
from flask_limiter.util import get_remote_address # Helper to get client IP for rate limiting

# --- Azure SDK Libraries ---
from azure.ai.formrecognizer import FormRecognizerClient # Client for Azure Form Recognizer service (AI-powered document analysis)
from azure.core.credentials import AzureKeyCredential    # For authenticating with Azure services using API keys
from azure.core.exceptions import HttpResponseError, ServiceRequestError, ResourceNotFoundError # Specific Azure SDK error types
from azure.core.polling import LROPoller                 # Type hint for long-running operation pollers (optional)
# --- Azure Storage ---
# Keep BlobServiceClient for potential future use or if other parts need it,
# but we won't use it for the intermediate steps in process_files
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient # Clients for interacting with Azure Blob Storage (though not used in the primary processing flow here)

# --- PDF Manipulation Libraries ---
from pypdf import PdfReader, PdfWriter   # For reading PDF metadata and merging pages (used in highlighting)
from reportlab.pdfgen import canvas      # For drawing shapes/text onto PDFs (used to create highlight overlays)
from reportlab.lib.units import inch     # Unit conversion for ReportLab coordinates

# -------------------- Flask App Setup --------------------
app = Flask(__name__, static_folder='static') # Initialize the Flask application
# Configure the maximum allowed size for uploads (e.g., 500 MB)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

# Initialize rate limiting based on remote IP address
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["200 per day", "50 per hour"] # Default limits for all routes unless overridden
)

# -------------------- Logging Setup ----------------------
# Configure basic logging to output informational messages and above
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(name)s - Thread: %(thread)d - %(message)s'
)
# Reduce log noise from verbose Azure libraries
logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
logging.getLogger("azure.storage.blob").setLevel(logging.WARNING)
logging.getLogger("pypdf").setLevel(logging.WARNING) # Quiet pypdf logs unless error
# Get a logger instance specifically for this application
logger = logging.getLogger(__name__)

# -------------------- Azure Form Recognizer Setup --------
# Retrieve Azure credentials from environment variables or hardcoded values
# WARNING: Hardcoded key! Acceptable ONLY for local testing as requested.
AZURE_FORM_RECOGNIZER_KEY = "KEY"
AZURE_FORM_RECOGNIZER_ENDPOINT = "Endpoint"

# Initialize the Form Recognizer client
try:
    azure_form_recognizer_client = FormRecognizerClient(
        AZURE_FORM_RECOGNIZER_ENDPOINT,
        AzureKeyCredential(AZURE_FORM_RECOGNIZER_KEY),
        logging_enable=False # Disable Azure SDK's verbose HTTP logging here if desired
    )
    logger.info("Azure Form Recognizer client initialized successfully.")
except Exception as e:
    logger.error(f"FATAL: Failed to initialize Azure Form Recognizer client: {e}", exc_info=True)
    azure_form_recognizer_client = None # Mark as unusable if initialization fails

# -------------------- Azure Storage Setup (Optional/Future Use) --------
# Retrieve Azure Storage connection string
AZURE_STORAGE_CONNECTION_STRING = "Connection_String"
AZURE_STORAGE_CONTAINER_NAME = "tempfiles" # Container name for potential future use (e.g., debugging, archival)

# Initialize the Blob Storage client (currently not used in the main file processing flow)
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

# -------------------- Utility Functions --------------------

def read_search_terms(excel_stream: io.BytesIO):
    """
    Reads search terms from the first column of an Excel file provided as a BytesIO stream.
    Performs case-insensitive matching by creating lowercase versions.

    Args:
        excel_stream (io.BytesIO): An in-memory byte stream containing the Excel file content.

    Returns:
        tuple: A tuple containing:
            - list: Original unique search terms.
            - set: Lowercase unique search terms for efficient lookup.
            - dict: Mapping from lowercase term back to its original casing.

    Raises:
        ValueError: If the Excel file is empty, contains no terms, or has format issues.
        RuntimeError: For unexpected errors during processing.
    """
    logger.info(f"Reading search terms from provided Excel stream.")
    try:
        excel_stream.seek(0) # Ensure stream pointer is at the beginning
        # Read only the first column (index 0), assume no header
        df = pd.read_excel(excel_stream, usecols=[0], header=None)
        # Extract terms, drop empty rows, convert to string, get unique values
        search_terms_original = df.iloc[:, 0].dropna().astype(str).unique().tolist()

        if not search_terms_original:
            raise ValueError("No search terms found in the first column of the Excel file.")

        # Prepare for case-insensitive matching
        search_terms_lower_set = set()
        lower_to_original_map = {}
        for term in search_terms_original:
            lower_term = term.lower()
            search_terms_lower_set.add(lower_term)
            # Store the first encountered original casing for each lower case term
            if lower_term not in lower_to_original_map:
                 lower_to_original_map[lower_term] = term

        logger.info(f"Read {len(search_terms_original)} unique search terms from stream.")
        return search_terms_original, search_terms_lower_set, lower_to_original_map
    except (ValueError) as e:
         # Log expected errors without stack trace for cleaner logs
         logger.error(f"Error reading Excel stream: {e}", exc_info=False)
         raise ValueError(f"Error reading or processing Excel file from stream: {e}")
    except Exception as e:
        # Log unexpected errors with stack trace
        logger.error(f"Unexpected error reading Excel stream: {e}", exc_info=True)
        raise RuntimeError(f"Unexpected error reading Excel file from stream: {e}")

def generate_excel_report(results):
    """
    Generates an Excel report summarizing the found keyword locations.

    Args:
        results (list): A list of dictionaries, where each dictionary represents a found keyword
                        and contains 'Keyword', 'Page', and 'WordObject' (with bounding_box).

    Returns:
        io.BytesIO: An in-memory byte stream containing the generated Excel report.

    Raises:
        RuntimeError: For unexpected errors during Excel generation.
    """
    if not results:
        # Create an empty DataFrame with correct headers if no results found
        df = pd.DataFrame(columns=["Keyword", "Page", "X1", "Y1", "X2", "Y2", "X3", "Y3", "X4", "Y4"])
    else:
        formatted_results = []
        # Iterate through results data provided by Azure Form Recognizer
        for item in results:
            # Extract bounding box coordinates from the WordObject
            coords = item['WordObject'].bounding_box
            if len(coords) == 4: # Ensure we have 4 points (corners) for the bounding box
                 formatted_results.append({
                    "Keyword": item['Keyword'], # The original keyword found
                    "Page": item['Page'],       # The page number where it was found
                    # Extract X and Y coordinates for each corner of the bounding box
                    "X1": coords[0].x, "Y1": coords[0].y, # Top-left
                    "X2": coords[1].x, "Y2": coords[1].y, # Top-right
                    "X3": coords[2].x, "Y3": coords[2].y, # Bottom-right
                    "X4": coords[3].x, "Y4": coords[3].y, # Bottom-left
                })
            else:
                # Log a warning if a bounding box doesn't have the expected 4 points
                 logger.warning(f"Keyword '{item['Keyword']}' on page {item['Page']} has an invalid bounding box length: {len(coords)}. Skipping this entry in the report.")

        # Create DataFrame, handling the case where all results had bad coordinates
        if not formatted_results:
            df = pd.DataFrame(columns=["Keyword", "Page", "X1", "Y1", "X2", "Y2", "X3", "Y3", "X4", "Y4"])
        else:
            df = pd.DataFrame(formatted_results)

    # Create an in-memory stream to hold the Excel data
    output = io.BytesIO()
    try:
        # Use pandas ExcelWriter to write the DataFrame to the stream
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Analysis")
        output.seek(0) # Reset stream position to the beginning for reading
        logger.info("Excel report generated successfully into memory stream.")
        return output
    except Exception as e:
        logger.error(f"Error generating Excel report: {e}", exc_info=True)
        raise RuntimeError(f"Error generating Excel report: {e}")

def process_pdf_azure(pdf_stream: io.BytesIO, search_terms_lower_set, lower_to_original_map, polling_interval_seconds=20):
    """
    Analyzes a PDF document from a BytesIO stream using Azure Form Recognizer's content analysis.
    Identifies occurrences of specified search terms (case-insensitive).

    Args:
        pdf_stream (io.BytesIO): In-memory byte stream of the PDF file.
        search_terms_lower_set (set): Set of lowercase keywords to search for.
        lower_to_original_map (dict): Mapping from lowercase keyword to its original casing.
        polling_interval_seconds (int): How often to check the status of the Azure job.

    Returns:
        list: A list of dictionaries, each containing 'Keyword', 'Page', and 'WordObject'
              for every matched keyword instance.

    Raises:
        ConnectionError: If the Form Recognizer client is unavailable or if Azure interaction fails (initial call or polling).
        RuntimeError: For unexpected errors or if the Azure job fails or finishes with an unknown status.
    """
    if not azure_form_recognizer_client:
        raise ConnectionError("Azure Form Recognizer client not available. Cannot process PDF.")

    try:
        logger.info(f"Starting Azure Form Recognizer content analysis on provided PDF stream.")
        pdf_stream.seek(0) # Ensure stream is at the beginning before sending to Azure

        # Start the asynchronous analysis process. This returns a poller object.
        poller = azure_form_recognizer_client.begin_recognize_content(
            pdf_stream,
            content_type='application/pdf', # Specify the content type
            logging_enable=False # Disable verbose Azure SDK logging for this call if desired
        )

        logger.info(f"Polling Azure for analysis results every {polling_interval_seconds} seconds...")
        start_poll_time = time.time()
        # Loop until the Azure job is complete (succeeded, failed, or cancelled)
        while not poller.done():
            try:
                current_status = poller.status()
                logger.debug(f"Polling... Current status: {current_status}")
                if poller.done():
                    break # Exit loop if already done
                # Wait for the specified interval before checking status again
                poller.wait(polling_interval_seconds)
            except TimeoutError:
                # Handle potential timeouts during the wait() call, just continue polling
                logger.debug(f"Polling wait timed out, checking status again...")
                continue
            except (HttpResponseError, ServiceRequestError) as poll_err:
                # Handle specific Azure errors during polling
                logger.error(f"Error during Azure polling: {poll_err}", exc_info=True)
                raise ConnectionError(f"Polling Azure status failed: {poll_err}") from poll_err
            except Exception as poll_err:
                 # Handle any other unexpected errors during polling
                 logger.error(f"Unexpected error during Azure polling wait: {poll_err}", exc_info=True)
                 raise RuntimeError(f"Polling failed unexpectedly: {poll_err}") from poll_err

        logger.info(f"Azure polling loop completed in {time.time() - start_poll_time:.2f} seconds.")

        # Check the final status of the Azure job
        final_status = poller.status()
        logger.info(f"Azure Form Recognizer job finished with status: {final_status}")

        # Process the results based on the final status
        if final_status.lower() == "succeeded":
            result = poller.result() # Get the analysis results
        elif final_status.lower() == "failed":
             try:
                 poller.result() # Calling result() on a failed job should raise an exception
             except HttpResponseError as e:
                 # Re-raise specific Azure error
                 raise ConnectionError(f"Azure analysis job failed (HTTP Status: {e.status_code}). Check Azure portal for details.") from e
             except Exception as e:
                 # Re-raise any other exception from result()
                 raise RuntimeError(f"Azure analysis job reported status '{final_status}' and failed: {e}") from e
             # Should not be reached if result() raises correctly
             raise RuntimeError(f"Azure analysis job reported status '{final_status}', but result() did not raise an error.")
        else:
             # Handle unexpected statuses
             raise RuntimeError(f"Azure analysis job finished with unexpected status: {final_status}")

        # --- Post-processing: Extract matched keywords from the results ---
        pdf_results = []
        # Regex to remove non-alphanumeric characters (excluding whitespace) for cleaner matching
        non_word_char_regex = re.compile(r'[^\w\s]')

        # Iterate through each page recognized by Form Recognizer
        for page in result:
            if page.lines: # Check if the page contains any lines of text
                # Iterate through lines and then words within each line
                for line in page.lines:
                    for word in line.words:
                        # Clean the recognized word: remove punctuation, convert to lowercase
                        cleaned_word = non_word_char_regex.sub('', word.text).lower()
                        # Check if the cleaned word matches any of the lowercase search terms
                        if cleaned_word in search_terms_lower_set:
                            # Retrieve the original casing of the keyword
                            original_term = lower_to_original_map.get(cleaned_word, word.text) # Fallback to raw text if somehow not in map
                            # Store the match details
                            pdf_results.append({
                                "Keyword": original_term,       # Original keyword text
                                "Page": page.page_number,       # Page number (1-based)
                                "WordObject": word              # The full Azure word object (contains text, bounding box, confidence)
                            })

        logger.info(f"Found {len(pdf_results)} keyword instances via Azure from PDF stream.")
        return pdf_results
    except (ConnectionError, RuntimeError) as e:
        # Re-raise known error types that signal processing failure
        raise e
    except Exception as e:
        # Catch any other unexpected errors during the process
        logger.error(f"Unexpected error during Azure PDF processing: {e}", exc_info=True)
        raise RuntimeError(f"An unexpected error occurred during Azure PDF analysis: {e}") from e


def highlight_pdf(original_pdf_stream: io.BytesIO, results_data):
    """
    Creates a new PDF with highlights around the keywords found by Form Recognizer.
    Uses pypdf to merge the original PDF content with highlight overlays created by reportlab.
    Operates entirely on in-memory BytesIO streams.

    Args:
        original_pdf_stream (io.BytesIO): In-memory byte stream of the original PDF.
        results_data (list): List of dictionaries from `process_pdf_azure`, containing
                             keyword locations ('Page', 'WordObject' with 'bounding_box').

    Returns:
        io.BytesIO: An in-memory byte stream containing the content of the new, highlighted PDF.
                    Returns a stream of the *original* PDF if highlighting fails or no results are provided.
    """
    if not results_data:
        logger.warning("No results data provided for highlighting. Returning original PDF content.")
        original_pdf_stream.seek(0) # Ensure stream is at the beginning
        # Return a *new* stream with the original content to avoid issues with closed streams elsewhere
        return io.BytesIO(original_pdf_stream.read())

    # --- Organize coordinates by page number ---
    page_coords = {} # Dictionary: {page_number: [list_of_coordinate_sets]}
    for item in results_data:
        page_idx = item['Page'] # 1-based page index from Form Recognizer
        word_obj = item['WordObject']
        coords_raw = word_obj.bounding_box # List of Point objects from Azure SDK

        # Validate coordinate data
        if len(coords_raw) == 4:
            try:
                # Convert Azure Point objects to (x, y) float tuples
                coords = [(float(p.x), float(p.y)) for p in coords_raw]
            except (ValueError, TypeError) as coord_err:
                logger.warning(f"Invalid coordinate value for word '{item['Keyword']}' on page {page_idx}. Skipping highlight. Error: {coord_err}")
                continue # Skip this specific highlight
        else:
            logger.warning(f"Word '{item['Keyword']}' on page {page_idx} has invalid bounding box length: {len(coords_raw)}. Skipping highlight.")
            continue # Skip this specific highlight

        # Store coordinates, grouping them by page number
        if page_idx not in page_coords:
            page_coords[page_idx] = []
        page_coords[page_idx].append(coords)

    # --- Create the highlighted PDF ---
    output_pdf_stream = io.BytesIO() # Stream to hold the final highlighted PDF
    try:
        original_pdf_stream.seek(0) # Reset original PDF stream position
        original_pdf_reader = PdfReader(original_pdf_stream)
        pdf_writer = PdfWriter() # Create a writer object for the new PDF
        num_pages = len(original_pdf_reader.pages)
        logger.info(f"Processing {num_pages} pages from PDF stream for highlighting (using add_blank/merge strategy).")

        # Iterate through each page of the original PDF
        for page_num in range(num_pages):
            page_index_1_based = page_num + 1 # Human-readable page number (1-based)
            original_page = original_pdf_reader.pages[page_num]

            # --- Create a new blank page and merge the original content onto it ---
            # This strategy avoids modifying the original page object directly and can be more robust.
            try:
                page_box = original_page.mediabox # Get dimensions (usually in points)
                pdf_width, pdf_height = float(page_box.width), float(page_box.height)
                # Add a blank page with the same dimensions as the original
                new_page = pdf_writer.add_blank_page(width=pdf_width, height=pdf_height)
                # Merge the content of the original page onto the new blank page
                new_page.merge_page(original_page)
                logger.debug(f"Merged original content for page {page_index_1_based}")
            except Exception as page_setup_err:
                 # Handle errors during page creation/merging (e.g., corrupted page)
                 logger.error(f"Error setting up or merging original page {page_index_1_based}: {page_setup_err}", exc_info=True)
                 try:
                     # Fallback: Add a standard US Letter size page if dimensions couldn't be read
                     pdf_writer.add_blank_page(width=612, height=792) # 8.5x11 inches in points
                     logger.warning(f"Added default blank page for page index {page_index_1_based} due to error.")
                 except Exception as add_blank_fallback_err:
                     # If even adding a blank page fails, log and skip this page
                     logger.error(f"Failed even to add a default blank page for page {page_index_1_based}: {add_blank_fallback_err}")
                 continue # Move to the next page

            # --- If there are highlights for this page, create and merge an overlay ---
            if page_index_1_based in page_coords:
                packet = io.BytesIO() # In-memory stream for the highlight overlay PDF
                try:
                    # Create a ReportLab canvas to draw on the overlay
                    # Use the dimensions obtained from the original page
                    c = canvas.Canvas(packet, pagesize=(pdf_width, pdf_height))
                    # Set drawing properties for the highlight box (red outline)
                    c.setStrokeColorRGB(1, 0, 0) # Red
                    c.setLineWidth(0.5)          # Thin line

                    # Draw each highlight box for the current page
                    for coords in page_coords[page_index_1_based]:
                        try:
                            p = c.beginPath() # Start defining a path
                            # ReportLab's origin (0,0) is bottom-left, while Azure's Y often increases downwards.
                            # Coordinates from Azure are typically in inches relative to top-left.
                            # Convert inches to points (ReportLab's default unit) and adjust Y coordinate.
                            pdf_height_points = pdf_height # Height in points
                            # MoveTo starts the path at the first corner (top-left)
                            p.moveTo(coords[0][0] * inch, pdf_height_points - coords[0][1] * inch)
                            # LineTo draws lines to subsequent corners in order (TL -> TR -> BR -> BL)
                            p.lineTo(coords[1][0] * inch, pdf_height_points - coords[1][1] * inch)
                            p.lineTo(coords[2][0] * inch, pdf_height_points - coords[2][1] * inch)
                            p.lineTo(coords[3][0] * inch, pdf_height_points - coords[3][1] * inch)
                            p.close() # Close the path to form a rectangle
                            # Draw the path outline (stroke=1), don't fill it (fill=0)
                            c.drawPath(p, stroke=1, fill=0)
                        except Exception as draw_err:
                            # Log if drawing a specific path fails, but continue with others
                            logger.warning(f"Error drawing highlight path on page {page_index_1_based} for coords {coords}. Skipping this highlight. Error: {draw_err}")
                            continue

                    c.save() # Finalize the overlay PDF content
                    packet.seek(0) # Reset the overlay stream position

                    # --- Merge the highlight overlay onto the new page ---
                    try:
                        overlay_reader = PdfReader(packet) # Read the generated overlay PDF
                        if overlay_reader.pages: # Ensure the overlay PDF isn't empty
                             overlay_page = overlay_reader.pages[0] # Get the first (only) page of the overlay
                             # Merge the overlay page onto the page containing the original content
                             new_page.merge_page(overlay_page)
                             logger.debug(f"Merged highlight overlay onto new page {page_index_1_based}")
                        else:
                             logger.warning(f"Highlight overlay generated for page {page_index_1_based} was empty or invalid. Skipping merge.")
                    except Exception as merge_err:
                        logger.error(f"Error merging highlight overlay onto new page {page_index_1_based}: {merge_err}", exc_info=True)
                        # Continue processing other pages even if one overlay merge fails
                finally:
                     # Ensure the temporary overlay stream is closed
                     packet.close()

        # --- Write the final PDF ---
        pdf_writer.write(output_pdf_stream) # Write all processed pages to the output stream
        output_pdf_stream.seek(0) # Reset stream position for sending
        logger.info(f"PDF highlighting completed successfully using pypdf/reportlab into memory stream.")
        return output_pdf_stream

    except (RuntimeError) as e:
        # Handle specific runtime errors that might be raised internally
        logger.error(f"Highlighting failed for PDF stream: {e}", exc_info=False)
        try:
            # Fallback: Return the original PDF content if highlighting fails
            logger.warning("Highlighting failed. Returning original PDF content.")
            original_pdf_stream.seek(0)
            return io.BytesIO(original_pdf_stream.read()) # Return a fresh stream
        except Exception as fallback_err:
             logger.error(f"Failed to return original PDF after highlighting error: {fallback_err}")
             # If even fallback fails, raise the original error
             raise RuntimeError(f"Highlighting failed, and could not retrieve original PDF: {e}") from e
    except Exception as e:
        # Handle any other unexpected errors during highlighting
        logger.error(f"Unexpected error during PDF highlighting: {e}", exc_info=True)
        # Raise a generic runtime error indicating highlight failure
        raise RuntimeError(f"An unexpected error occurred during PDF highlighting: {e}")


# -------------------- Flask Routes --------------------

@app.route("/", methods=["GET"])
def index():
    """Serves the main HTML page of the application."""
    logger.info(f"Serving index.html to {get_remote_address()}")
    return render_template("index.html")

@app.route("/process", methods=["POST"])
@limiter.limit("5/minute") # Apply specific rate limit to this potentially heavy endpoint
def process_files():
    """
    Handles the file upload, processing, and result generation.
    Orchestrates the calls to utility functions for PDF analysis and highlighting.
    Returns a ZIP file containing the analysis report (Excel) and the highlighted PDF.
    """
    start_time = time.time()
    request_ip = get_remote_address()
    logger.info(f"Received /process request from {request_ip}")

    # --- Check Service Availability ---
    if not azure_form_recognizer_client:
        # If the Form Recognizer client failed to initialize, return Service Unavailable
        logger.error(f"Rejecting /process request from {request_ip} due to unavailable Form Recognizer client.")
        abort(503, description="Server error: The document analysis service is currently unavailable. Please try again later.")

    # --- Get Uploaded Files ---
    pdf_file = request.files.get("pdf_file")
    excel_file = request.files.get("excel_file")

    # --- Basic Input Validation ---
    if not pdf_file or not excel_file or not pdf_file.filename or not excel_file.filename:
        logger.warning(f"Missing files or filenames in request from {request_ip}.")
        abort(400, description="Missing required PDF or Excel file, or filename is empty.")

    # Sanitize filenames to prevent directory traversal or other injection attacks
    safe_pdf_filename = secure_filename(pdf_file.filename)
    safe_excel_filename = secure_filename(excel_file.filename)
    # Extract base name for use in output filenames
    pdf_base, _ = os.path.splitext(safe_pdf_filename)
    # Generate a unique ID for this processing run for logging/tracking
    run_id = str(uuid.uuid4())

    # --- MIME Type and Extension Validation ---
    allowed_pdf_mimetypes = ["application/pdf"]
    allowed_excel_mimetypes = ["application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
    allowed_excel_extensions = [".xls", ".xlsx"]

    pdf_mimetype = pdf_file.mimetype
    excel_mimetype = excel_file.mimetype
    _, pdf_ext = os.path.splitext(safe_pdf_filename) # Check extension as fallback
    _, excel_ext = os.path.splitext(safe_excel_filename) # Check extension as fallback

    if pdf_mimetype not in allowed_pdf_mimetypes:
        logger.warning(f"Invalid PDF file type '{pdf_mimetype}' from {request_ip} (Run ID: {run_id}).")
        abort(400, description=f"Invalid file type for Lab Notebook. Expected PDF (application/pdf), got '{pdf_mimetype}'.")
    # Check both MIME type and extension for Excel as MIME types can be inconsistent
    if excel_mimetype not in allowed_excel_mimetypes and excel_ext.lower() not in allowed_excel_extensions:
        logger.warning(f"Invalid Excel file type '{excel_mimetype}' / ext '{excel_ext}' from {request_ip} (Run ID: {run_id}).")
        abort(400, description=f"Invalid file type for Keywords. Expected Excel (.xls or .xlsx), got type '{excel_mimetype}' / extension '{excel_ext}'.")

    logger.info(f"Processing run {run_id}: PDF='{safe_pdf_filename}', Excel='{safe_excel_filename}' from {request_ip}")

    # Initialize variables for in-memory streams
    pdf_stream_in_memory = None
    excel_stream_in_memory = None
    excel_report_stream = None
    highlighted_pdf_stream = None
    output_zip_stream = None

    try:
        # --- Step 1: Read input files fully into memory streams ---
        # This avoids saving temporary files to disk, suitable for moderate file sizes.
        pdf_file.seek(0) # Ensure pointer is at the start of the file stream
        pdf_stream_in_memory = io.BytesIO(pdf_file.read())
        excel_file.seek(0) # Ensure pointer is at the start of the file stream
        excel_stream_in_memory = io.BytesIO(excel_file.read())
        logger.info(f"Run {run_id}: Input files read into memory streams.")

        # --- Step 2: Read Search Terms (from the Excel memory stream) ---
        try:
            # Call the utility function to parse keywords from the Excel stream
            _, search_terms_lower_set, lower_to_original_map = read_search_terms(excel_stream_in_memory)
        except (ValueError, RuntimeError) as e:
            # Handle errors during keyword reading (e.g., bad format, empty file)
            logger.warning(f"Run {run_id}: Failed to read search terms from Excel stream: {e}", exc_info=False)
            # Determine appropriate HTTP status code based on error type
            status_code = 400 if isinstance(e, ValueError) else 500
            abort(status_code, description=str(e)) # Send error description to client

        # --- Step 3: Process PDF with Azure Form Recognizer (using the PDF memory stream) ---
        try:
            # Call the utility function to analyze the PDF content using Azure
            results = process_pdf_azure(pdf_stream_in_memory, search_terms_lower_set, lower_to_original_map)
        except (ConnectionError, RuntimeError) as e:
            # Handle errors during Azure processing (connection issues, Azure job failure)
            logger.error(f"Run {run_id}: Azure PDF processing failed from stream: {e}", exc_info=False) # Log without full stack trace for cleaner logs unless debugging
            status_code = 503 if isinstance(e, ConnectionError) else 500 # 503 for Azure connection issues, 500 otherwise
            abort(status_code, description=f"Failed to analyze PDF: {e} (Run ID: {run_id})") # Include Run ID in error message
        except Exception as e:
             # Catch any other unexpected errors during PDF processing
             logger.error(f"Run {run_id}: Unexpected error during PDF processing from stream: {e}", exc_info=True)
             abort(500, description=f"An unexpected error occurred during PDF analysis. (Run ID: {run_id})")

        # --- Step 4: Generate Excel Report (results into a new memory stream) ---
        try:
            # Call the utility function to create the Excel report based on Azure results
            excel_report_stream = generate_excel_report(results)
        except (RuntimeError, Exception) as e:
             # Handle errors during Excel report generation
             logger.error(f"Run {run_id}: Excel report generation failed: {e}", exc_info=True)
             abort(500, description=f"Failed to generate the analysis report. (Run ID: {run_id})")

        # --- Step 5: Highlight PDF (using original PDF stream and results, output to new memory stream) ---
        try:
            # Call the utility function to create the highlighted PDF
            # Pass the *original* PDF stream again, as previous steps might have consumed it
            highlighted_pdf_stream = highlight_pdf(pdf_stream_in_memory, results)
        except (RuntimeError, Exception) as e:
             # Handle errors during PDF highlighting
             logger.error(f"Run {run_id}: PDF highlighting failed: {e}", exc_info=True)
             abort(500, description=f"Failed to create the highlighted PDF. (Run ID: {run_id})")

        # --- Step 6: Create ZIP archive in memory ---
        output_zip_stream = io.BytesIO() # Create a stream for the ZIP file
        try:
            # Reset positions of the streams containing the report and highlighted PDF
            excel_report_stream.seek(0)
            highlighted_pdf_stream.seek(0)
            # Create a ZipFile object writing to the output stream
            with zipfile.ZipFile(output_zip_stream, "w", zipfile.ZIP_DEFLATED) as zf:
                # Add the Excel report to the ZIP
                zf.writestr(f"{pdf_base}_analysis.xlsx", excel_report_stream.getvalue())
                # Add the highlighted PDF to the ZIP
                zf.writestr(f"{pdf_base}_highlighted.pdf", highlighted_pdf_stream.getvalue())
            output_zip_stream.seek(0) # Reset ZIP stream position for sending
            logger.info(f"Run {run_id}: Successfully created results ZIP file in memory.")
        except (zipfile.BadZipFile, Exception) as zip_err:
             # Handle errors during ZIP creation
             logger.error(f"Run {run_id}: Failed to create ZIP file in memory: {zip_err}", exc_info=True)
             abort(500, description=f"Failed to create the final results ZIP file. (Run ID: {run_id})")

        # --- Step 7: Send the ZIP file as a download ---
        zip_download_filename = f"{pdf_base}_results.zip" # Define the filename for the download
        total_time = time.time() - start_time
        logger.info(f"Run {run_id}: Successfully processed '{safe_pdf_filename}'. Sending ZIP '{zip_download_filename}'. Total Time: {total_time:.2f}s IP: {request_ip}.")
        # Use Flask's send_file to send the in-memory ZIP stream
        return send_file(
            output_zip_stream,
            mimetype="application/zip",      # Set the correct MIME type for ZIP files
            as_attachment=True,              # Trigger browser download dialog
            download_name=zip_download_filename # Specify the filename for the download
        )

    except Exception as e:
        # --- Generic Exception Handler for the entire route ---
        # Log any uncaught exceptions that occurred during the request processing
        logger.exception(f"Run {run_id}: Unhandled exception in /process route for IP {request_ip}:") # logger.exception includes stack trace
        # Check if response headers have already been sent (e.g., if error happens after send_file starts)
        if not request.environ.get('werkzeug.response_started'):
            # If response hasn't started, return a generic 500 error to the client
            return jsonify({"error": f"An unexpected server error occurred. Please contact support. (Ref: {run_id})"}), 500
        else:
             # If response already started, we can't send a new JSON response. Log the error.
             logger.error(f"Run {run_id}: Exception occurred after response headers were sent. Cannot send error JSON.")
             # Return None or let the server handle the broken connection
             return None

    finally:
        # --- Clean up in-memory streams ---
        # Ensure streams are closed to free up memory, regardless of success or failure
        if pdf_stream_in_memory: pdf_stream_in_memory.close()
        if excel_stream_in_memory: excel_stream_in_memory.close()
        if excel_report_stream: excel_report_stream.close()
        if highlighted_pdf_stream: highlighted_pdf_stream.close()
        # DO NOT CLOSE output_zip_stream here. Flask's send_file (or the WSGI server like Waitress)
        # is responsible for closing the stream after sending the response. Closing it here would cause errors.

# -------------------- Error Handlers --------------------
# Define custom handlers for specific HTTP error codes to provide consistent JSON responses.

@app.errorhandler(400)
def bad_request(error):
    """Handles HTTP 400 Bad Request errors (e.g., missing files, invalid types)."""
    description = getattr(error, 'description', "Invalid request parameters.")
    logger.warning(f"Bad Request (400) from {get_remote_address()}: {description}")
    return jsonify({"error": description}), 400

@app.errorhandler(413)
def request_entity_too_large(error):
    """Handles HTTP 413 Payload Too Large errors (file size exceeds MAX_CONTENT_LENGTH)."""
    max_size_mb = app.config["MAX_CONTENT_LENGTH"] / (1024 * 1024)
    description = getattr(error, 'description', f"File size exceeds the server limit of {max_size_mb:.1f} MB.")
    logger.warning(f"Request Entity Too Large (413) from {get_remote_address()}. Limit: {max_size_mb:.1f} MB")
    return jsonify(error=description), 413

@app.errorhandler(429)
def ratelimit_handler(error):
    """Handles HTTP 429 Too Many Requests errors (rate limit exceeded)."""
    # The description is automatically set by Flask-Limiter
    limit_info = getattr(error, 'description', "Rate limit exceeded")
    logger.warning(f"Rate Limit Exceeded (429) for {get_remote_address()}. Limit info: {limit_info}")
    # Provide a user-friendly message
    return jsonify(error=f"You have exceeded the request limit ({limit_info}). Please try again later."), 429

@app.errorhandler(500)
def internal_server_error(error):
    """Handles HTTP 500 Internal Server Error (uncaught exceptions)."""
    run_id_str = ""
    # Attempt to extract the Run ID from the error description if it was added via abort()
    description = getattr(error, 'description', "An internal server error occurred.")
    if isinstance(description, str):
        match = re.search(r'\((?:Ref|Run ID):\s*([a-fA-F0-9\-]+)\)', description)
        if match:
            run_id_str = f" (Ref: {match.group(1)})"

    # Log the error with details
    original_exception = getattr(error, 'original_exception', error) # Get original exception if available
    logger.error(f"Internal Server Error (500) from {get_remote_address()}{run_id_str}: {description}", exc_info=original_exception)
    # Provide a generic error message to the user, potentially including the reference ID
    user_message = f"An unexpected internal server error occurred{run_id_str}. Please contact support if the issue persists."
    return jsonify(error=user_message), 500

@app.errorhandler(503)
def service_unavailable(error):
     """Handles HTTP 503 Service Unavailable errors (e.g., dependent service like Azure is down)."""
     description = getattr(error, 'description', "The service is temporarily unavailable.")
     logger.error(f"Service Unavailable (503) from {get_remote_address()}: {description}")
     return jsonify(error=description), 503

# -------------------- Main Execution --------------------
if __name__ == "__main__":
    # --- Pre-run Checks ---
    # Ensure the static folder exists (where index.html might look for CSS/JS/images)
    static_folder_path = os.path.join(os.path.dirname(__file__), 'static')
    if not os.path.exists(static_folder_path):
        try:
            os.makedirs(static_folder_path)
            logger.info(f"Created 'static' folder at: {static_folder_path}")
        except OSError as e:
            logger.error(f"Failed to create 'static' folder: {e}")

    # Check status of critical Azure clients initialized earlier
    clients_ok = True
    if azure_form_recognizer_client is None:
        print("\n!! FATAL ERROR: Azure Form Recognizer client failed to initialize. Application cannot function correctly. !!")
        clients_ok = False
    if azure_blob_service_client is None:
        # This is currently less critical as blob storage isn't in the main path
        print("\n!! WARNING: Azure Blob Storage client failed initialization. This may affect optional features but core processing might still work. !!")

    if not clients_ok:
        print("\n!! Critical Azure client initialization failed. Exiting. Check logs for details. !!\n")
        import sys; sys.exit(1) # Exit if critical components are missing
    else:
         # Print client status if initialization was okay
         print("\n--- Azure Clients Initialized ---")
         print(f"Form Recognizer Endpoint: {AZURE_FORM_RECOGNIZER_ENDPOINT}")
         if azure_blob_service_client:
             print(f"Storage Account: {azure_blob_service_client.account_name} / Container: {AZURE_STORAGE_CONTAINER_NAME} (Initialized, but may not be used in critical path)")
         else:
             print("Storage Account: Client failed initialization (not critical for current flow).")
         print("---------------------------------\n")

    # --- Server Configuration ---
    # Get port from environment variable 'PORT', common practice for PaaS like Azure Web Apps.
    # Default to 5000 for local development if PORT is not set.
    port = int(os.environ.get("PORT", 5000))

    # --- Start Server ---
    print(f"Starting Flask application server using Waitress on http://0.0.0.0:{port}")
    print(f"Serving static files from: {app.static_folder}")
    print("Ensure 'epiqlogo.png' is present in the 'static' directory for the UI.")
    print("\n*** DEVELOPMENT WARNING: Review Azure credential handling before deploying to production. ***\n")
    print("*** Using in-memory processing for intermediate files. Monitor memory usage for large files. ***\n")

    # Use Waitress, a production-quality WSGI server, instead of Flask's built-in development server
    from waitress import serve
    serve(
        app,
        host='0.0.0.0',      # Listen on all available network interfaces
        port=port,           # Use the configured port
        threads=8,           # Number of worker threads to handle requests concurrently
        channel_timeout=900  # Increase timeout for potentially long Azure operations (15 minutes)
    )
