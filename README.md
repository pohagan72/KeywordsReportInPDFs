# Keywords Report In PDFs

This Flask web application automates the process of identifying and highlighting specific keywords within PDF files. It leverages Azure Form Recognizer to extract text from PDFs, searches for user-defined keywords from an Excel file, and generates both Excel reports and highlighted PDFs indicating where the keywords were found.

## Key Features

1. **Web Interface**: User-friendly web interface with drag-and-drop functionality for file uploads
2. **PDF Text Extraction**: Utilizes Azure Form Recognizer Layout API to extract text content and bounding boxes
3. **Keyword Search**: 
   - Case-insensitive search with punctuation removal for robust matching
   - Preserves original keyword casing in reports
4. **Reporting**:
   - Detailed Excel reports with keyword locations (page number and coordinates)
   - Comprehensive error handling and logging
5. **PDF Highlighting**: 
   - Creates highlighted versions of original PDFs
   - Maintains original document quality
6. **Security**:
   - Rate limiting to prevent abuse
   - Secure credential handling via environment variables
7. **Progress Tracking**: Real-time progress indication during processing

## How It Works

1. **User Input**:
   - Upload a PDF file through the web interface
   - Upload an Excel file containing keywords (first column is read as keywords)

2. **Processing**:
   - PDF is sent to Azure Form Recognizer for text and layout analysis
   - Extracted text is searched for keyword matches
   - Results are compiled with page numbers and coordinates
   - Original PDF is annotated with highlights at keyword locations

3. **Output**:
   - Single ZIP file containing:
     - Highlighted PDF
     - Excel report with all keyword matches

## Setup and Configuration

### Prerequisites

- Python 3.8+
- Azure account with Form Recognizer service
- Azure Storage account (optional, for handling large files)

### Installation

1. Clone the repository:
   ```bash
   git clone [repository-url]
   cd KeywordsReportInPDFs
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure environment variables:
   ```bash
   export AZURE_FORM_RECOGNIZER_ENDPOINT="your-azure-endpoint"
   export AZURE_FORM_RECOGNIZER_KEY="your-azure-key"
   export AZURE_STORAGE_CONNECTION_STRING="your-storage-connection-string"  # Optional
   ```

### Running the Application

Start the development server:
```bash
python main.py
```

For production use:
```bash
gunicorn -w 4 -b :5000 main:app
```

The application will be available at `http://localhost:5000`

## Technical Details

### Architecture

- **Frontend**: HTML/JS interface with drag-and-drop functionality
- **Backend**: Flask application with:
  - Azure Form Recognizer integration
  - PDF processing using PyPDF and ReportLab
  - Excel report generation with pandas
- **Error Handling**: Comprehensive error handling and logging

### File Processing Flow

1. User uploads files via web interface
2. Server validates file types and sizes
3. Excel keywords are extracted and normalized
4. PDF is processed through Azure Form Recognizer
5. Keyword matches are identified and recorded
6. Highlighted PDF is generated
7. Excel report is created
8. Results are packaged into ZIP file
9. ZIP file is returned to user

## Security Considerations

- Credentials are never hardcoded (use environment variables)
- Rate limiting prevents abuse (200 requests/day, 50 requests/hour by default)
- File uploads are validated for type and size
- Comprehensive error handling prevents information leakage

## Limitations

- Currently processes one PDF at a time
- Maximum upload size: 500MB (configurable)
- Requires Azure Form Recognizer service
- Case-insensitive matching only

## Future Enhancements

- Batch processing of multiple PDFs
- Configurable matching options (case sensitivity, partial matches)
- Support for additional file formats
- Asynchronous processing for large documents
