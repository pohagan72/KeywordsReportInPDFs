# Keywords Report In PDFs

A Flask-based web application designed to analyze Lab Notebooks by extracting text with Azure Form Recognizer, identifying user-specified keywords from an Excel file, and generating two outputs:
- An **Excel report** detailing the occurrences and locations (page number and bounding box coordinates) of each keyword.
- A **highlighted PDF** where the found keywords are visually marked.

The two output files are packaged together into a ZIP file and sent to the user for download.

---

## Table of Contents

- [Features](#features)
- [Architecture Overview](#architecture-overview)
- [Setup and Installation](#setup-and-installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [File Structure](#file-structure)
- [Security Considerations](#security-considerations)
- [Future Enhancements](#future-enhancements)

---

## Features

- **PDF Analysis:**  
  Uses the Azure Form Recognizer Layout API to extract text and layout information (bounding boxes) from uploaded PDFs.

- **Keyword Matching:**  
  Reads keywords from an Excel file (first column) and performs case-insensitive matching while cleaning punctuation for accurate search results.

- **Excel Report Generation:**  
  Creates a detailed Excel report using pandas. The report includes each keyword found, the corresponding page number, and its bounding box coordinates.

- **PDF Highlighting:**  
  Generates a new PDF by overlaying highlight boxes on the original document using pypdf and ReportLab. If highlighting fails, the original PDF is returned.

- **User-Friendly Web Interface:**  
  The HTML interface offers a drag-and-drop experience for uploading files (PDF for lab notebooks and Excel for keywords), along with real-time progress and error feedback.

- **Rate Limiting:**  
  Protects the service using Flask-Limiter to avoid abuse (default limits set to 200 requests per day and 50 requests per hour overall, with 5 requests per minute on the processing endpoint).

- **Error Handling and Logging:**  
  Comprehensive error management throughout the processing flow, including handling Azure-related issues, file validation errors, and unexpected exceptions.

---

## Architecture Overview

1. **Frontend (HTML/JavaScript):**  
   - **Template:** `templates/index.html` provides an intuitive UI with drag-and-drop zones for file uploads.
   - **Interactivity:** JavaScript handles file selection updates, form submission, and displays a loading spinner during processing.

2. **Backend (Flask Application):**  
   - **Core Application:** Defined in `main.py`, it orchestrates the complete file processing workflow:
     - **Input Validation:** Verifies the uploaded file types and ensures filenames are secure.
     - **Keyword Extraction:** Reads keywords from the provided Excel file.
     - **PDF Processing:** Uses Azure Form Recognizer to extract text and layout information from the PDF.
     - **Report and Highlight Generation:** Creates an Excel report and a highlighted version of the PDF.
     - **Output Packaging:** Bundles the Excel and PDF files into a ZIP archive before returning them to the user.
   - **Azure Integration:**  
     - Utilizes the Azure Form Recognizer client for document analysis.
     - (Optional/Future) Integration with Azure Blob Storage is in place but not part of the main processing flow.
   - **Production Server:**  
     - Uses [Waitress](https://docs.pylonsproject.org/projects/waitress/en/stable/) as a production-quality WSGI server.

---

## Setup and Installation

### Prerequisites

- **Python:** Version 3.8 or higher.
- **Azure Account:** An active Azure account with access to the Form Recognizer service.
- **Optional (Azure Blob Storage):** For future enhancements and handling large files.

### Installation Steps

1. **Clone the Repository:**

   ```bash
   git clone [repository-url]
   cd KeywordsReportInPDFs
   ```

2. **Create a Virtual Environment (Recommended):**

   ```bash
   python -m venv venv
   source venv/bin/activate   # On Windows use: venv\Scripts\activate
   ```

3. **Install Dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

4. **Configure Environment Variables:**

   Set the following environment variables before starting the application:

   ```bash
   export AZURE_FORM_RECOGNIZER_ENDPOINT="your-azure-endpoint"
   export AZURE_FORM_RECOGNIZER_KEY="your-azure-key"
   export AZURE_STORAGE_CONNECTION_STRING="your-storage-connection-string"  # Optional if needed
   ```

   You can also create a `.env` file and use a package like `python-dotenv` to load these variables.

---

## Configuration

- **Maximum Upload Size:**  
  The application is configured to accept uploads up to 500 MB.

- **Rate Limiting:**  
  Default rate limits are applied for the entire application and specifically for the file processing endpoint (`/process`).

- **Logging:**  
  Basic logging is set up to capture INFO-level messages and above. Adjust the logging configuration in `main.py` as needed.

- **Azure Clients:**  
  The application initializes the Azure Form Recognizer and Azure Blob Storage clients on startup. If the Form Recognizer client fails to initialize, the app will exit with a critical error.

---

## Usage

1. **Start the Application:**

   For development purposes, you can run:

   ```bash
   python main.py
   ```

   For production, consider using Waitress or Gunicorn:

   ```bash
   waitress-serve --host=0.0.0.0 --port=5000 main:app
   ```

2. **Access the Web Interface:**

   Open your browser and navigate to `http://localhost:5000`.

3. **Upload Files:**

   - **Lab Notebook (PDF):** Drag and drop or select a PDF file.
   - **Keywords (Excel):** Drag and drop or select an Excel file containing keywords in the first column.

4. **Submit and Download:**

   Click the **Analyze Files** button. Once processed, a ZIP file containing the highlighted PDF and Excel report will be downloaded automatically.

---

## File Structure

```
KeywordsReportInPDFs/
├── main.py                    # Main Flask application with file processing logic
├── templates/
│   └── index.html             # HTML template for the front-end drag-and-drop interface
├── static/
│   └── epiqlogo.png           # Logo used in the interface (ensure this file is present)
├── README.md                  # This file
└── requirements.txt           # Python dependencies (Flask, azure-ai-formrecognizer, pandas, pypdf, reportlab, etc.)
```

---

## Future Enhancements

- **Batch Processing:**  
  Support for processing multiple PDFs in a single request.

- **Improved Keyword Matching:**  
  Additional matching options (e.g., partial matches, configurable case sensitivity).

- **Enhanced Azure Storage Integration:**  
  Utilize Azure Blob Storage for temporary file storage, logging, or archival.

- **UI/UX Improvements:**  
  Further enhancements to the user interface and progress indicators during processing.
---
