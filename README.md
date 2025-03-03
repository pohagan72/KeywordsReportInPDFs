# Keywords Report In PDFs

This application automates the process of identifying and highlighting specific keywords within a directory of PDF files. It leverages Azure Form Recognizer to extract text from PDFs, searches for user-defined keywords, and generates both Excel reports and highlighted PDFs indicating where the keywords were found.

## Functionality

1.  **PDF Text Extraction:** Utilizes Azure Form Recognizer to extract text content and layout information from PDF documents.

2.  **Keyword Search:** Searches for exact matches of user-provided keywords within the extracted text. Case-insensitive search is performed to ensure comprehensive results.

3.  **Reporting:**
    *   Generates Excel reports detailing the keywords found, the page numbers on which they appear, and their coordinates within each PDF.
    *   Creates a combined Excel report that summarizes the findings across all processed PDF files.

4.  **PDF Highlighting:** Highlights the identified keywords directly within the PDF files, making it easy to visually locate the search terms.

5.  **Zipped Output:** Zips all highlighted PDFs into a single downloadable archive for convenient access.

## How It Works

1.  **User Input:**
    *   The user provides a folder path containing the PDF files to be analyzed.
    *   The user uploads an Excel file containing a list of keywords to search for.  The first column of the excel file is read and the program takes each row as a keyword.

2.  **Processing:**

    *   The application iterates through each PDF file in the specified folder.
    *   For each PDF:
        *   The PDF content is sent to Azure Form Recognizer for text extraction.
        *   The extracted text is searched for the keywords specified in the Excel file.
        *   An Excel report is generated, listing each found keyword, its page number, and coordinates.
        *   A highlighted PDF is created, with the found keywords highlighted.
    *   A combined Excel report is generated, summarizing the findings across all PDFs.
    *   All highlighted PDFs are zipped into a single archive.

3.  **Output:**

    *   The user can download the zip archive containing all the highlighted PDFs.
    *   Excel reports are generated for each PDF and a combined excel report is created.

## Setup and Configuration

1.  **Azure Form Recognizer:**

    *   You need an Azure subscription and access to Azure Form Recognizer.
    *   Obtain your Azure Form Recognizer key and endpoint URL.
    *   Replace the placeholder values in `main.py`:

        ```python
        AZURE_KEY = "YOUR_AZURE_KEY"  # Replace with your key
        ENDPOINT = "YOUR_AZURE_ENDPOINT" # Replace with your endpoint
        ```

2.  **Dependencies:**

    *   Install the required Python packages:

        ```bash
        pip install pandas azure-ai-formrecognizer reportlab PyPDF2 flask
        ```

3.  **Folder Structure:**

    *   Create a folder to store your PDF files.
    *   Prepare an Excel file with the list of keywords to search for (one keyword per row in the first column).

## Usage

1.  **Run the Application:**

    ```bash
    python main.py
    ```

2.  **Access the Web Interface:**

    *   Open your web browser and navigate to `http://127.0.0.1:5000/` (or the address shown in the console when you run the app).

3.  **Input:**

    *   Enter the path to the folder containing your PDF files.
    *   Upload the Excel file with the keywords.

4.  **Process:**

    *   Click the "Start Processing" button.

5.  **Download:**

    *   After processing is complete, the highlighted PDFs will automatically download as a zip file.

## Notes

*   The application uses exact matching for keyword search (case-insensitive).
*   Ensure that the Azure Form Recognizer key and endpoint are correctly configured.
*   The Excel file should have the keywords listed in the first column.
*   The application provides a progress bar to track the processing status.
