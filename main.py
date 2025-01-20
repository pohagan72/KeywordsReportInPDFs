# Import necessary libraries
import streamlit as st  # For building the web app interface
import os  # For interacting with the file system
import time  # For tracking processing time
import pandas as pd  # For handling Excel files and data manipulation
import warnings  # For suppressing warnings
from azure.ai.formrecognizer import FormRecognizerClient  # For using Azure Form Recognizer
from azure.core.credentials import AzureKeyCredential  # For authentication with Azure

# Suppress warnings related to the openpyxl library (used by pandas for Excel files)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Azure Form Recognizer credentials (replace with your actual credentials if needed)
AZURE_KEY = "Your_Key_Goes_Here"
ENDPOINT = "Your_Endpoint_Goes_Here"

# Initialize the Azure Form Recognizer client
try:
    # This client will be used to interact with the Azure Form Recognizer service
    form_recognizer_client = FormRecognizerClient(ENDPOINT, AzureKeyCredential(AZURE_KEY))
except Exception as e:
    # If initialization fails, show an error message and stop the Streamlit app
    st.error(f"Failed to initialize Form Recognizer Client: {e}")
    st.stop()

def process_pdf(pdf_path: str, search_terms: list, client: FormRecognizerClient):
    """
    Processes a single PDF file using Azure Form Recognizer to find pages containing specified search terms.

    :param pdf_path: Path to the PDF file.
    :param search_terms: A list of terms to search for in the document.
    :param client: An initialized FormRecognizerClient instance.
    :return: A list of dictionaries containing page numbers and found search terms.
    """
    pdf_results = []  # List to store results for each page where search terms are found

    try:
        # Open the PDF file in binary mode and send it to Azure Form Recognizer
        with open(pdf_path, "rb") as f:
            poller = client.begin_recognize_content(f)  # Start the recognition process
            recognized_content = poller.result()  # Retrieve the results
    except Exception as e:
        # If there's an issue during recognition, log the error and return an empty list
        st.error(f"Error processing {os.path.basename(pdf_path)}: {e}")
        return pdf_results

    # Check if any content was recognized in the PDF
    if not recognized_content:
        st.warning(f"No pages recognized in {os.path.basename(pdf_path)}. Skipping.")
        return pdf_results

    # Iterate over each recognized page in the PDF
    for i, page in enumerate(recognized_content):
        try:
            # Extract lines of text from the page and convert them to lowercase for case-insensitive search
            page_text = [line.text.lower() for line in page.lines]
            # Split lines into individual words for easier search
            words = [word for line in page_text for word in line.split()]
        except AttributeError:
            # If the page does not have any lines, skip it
            continue

        # Search for the specified terms in the current page's text
        found_terms = [term for term in search_terms if term.lower() in words]
        if found_terms:
            # If search terms are found, create a dictionary with the page number and found terms
            row_data = {"Page": i + 1}  # Page numbers are 1-based
            for col_index, term in enumerate(found_terms, start=1):
                row_data[f"Keyword {col_index}"] = term  # Associate each found term with a column
            pdf_results.append(row_data)  # Add the result to the list

    return pdf_results

# Streamlit App UI
st.title("Batch PDF Word Search with Azure Form Recognizer")  # App title
st.write("Enter a folder path containing PDFs and upload an Excel file with search terms.")  # Instructions

# Input field for the folder containing PDFs
folder_path = st.text_input("Enter folder path containing PDFs:")

# File uploader for the Excel file containing search terms
uploaded_excel = st.file_uploader("Upload Excel file with search terms", type=["xlsx", "xls"])

# Button to start processing
if st.button("Start Processing") and folder_path and uploaded_excel:
    # Validate the provided folder path
    if not os.path.isdir(folder_path):
        st.error("Invalid folder path. Please enter a valid directory.")
        st.stop()

    try:
        # Read the search terms from the uploaded Excel file
        df_terms = pd.read_excel(uploaded_excel)
        search_terms = df_terms.iloc[:, 0].dropna().tolist()  # Extract the first column (non-empty values)
    except Exception as e:
        # If reading the Excel file fails, show an error message and stop the app
        st.error(f"Could not read the Excel file: {e}")
        st.stop()

    # Ensure that search terms were successfully extracted
    if not search_terms:
        st.error("No search terms found in the Excel file.")
        st.stop()

    st.write(f"Loaded {len(search_terms)} search terms.")  # Display the number of search terms loaded

    # Get the list of PDF files in the specified folder
    try:
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    except Exception as e:
        st.error(f"Error listing PDF files: {e}")
        st.stop()

    # Check if any PDFs were found
    if not pdf_files:
        st.warning("No PDF files found in the specified folder.")
        st.stop()

    total_pdfs = len(pdf_files)  # Total number of PDFs to process
    st.write(f"Found {total_pdfs} PDF file(s) to process.")  # Display the number of PDFs found

    # Initialize a progress bar and placeholders for status and timing updates
    progress_bar = st.progress(0)
    status_text = st.empty()
    time_text = st.empty()

    # Record the start time of the entire processing operation
    overall_start_time = time.time()

    # Process each PDF file one by one
    for pdf_index, pdf_file in enumerate(pdf_files, start=1):
        pdf_path = os.path.join(folder_path, pdf_file)  # Get the full path to the current PDF

        pdf_start_time = time.time()  # Record the start time for this PDF

        # Process the PDF to extract pages containing search terms
        pdf_results = process_pdf(pdf_path, search_terms, form_recognizer_client)

        # If results were found, save them to an Excel file
        if pdf_results:
            output_filename = os.path.join(
                folder_path, f"{os.path.splitext(pdf_file)[0]}_Analysis.xlsx"
            )
            try:
                df_results = pd.DataFrame(pdf_results)  # Create a DataFrame from the results
                df_results.to_excel(output_filename, index=False)  # Save the results to an Excel file
            except Exception as e:
                st.error(f"Error saving results for {pdf_file}: {e}")

        # Update the progress bar and status information
        current_time = time.time()
        elapsed_time = current_time - overall_start_time  # Total time elapsed
        avg_time_per_pdf = elapsed_time / pdf_index if pdf_index > 0 else 0  # Average time per PDF

        progress_percent = int((pdf_index / total_pdfs) * 100)  # Calculate progress percentage
        progress_bar.progress(progress_percent)  # Update progress bar
        status_text.text(f"Processing PDF {pdf_index} of {total_pdfs}: {pdf_file}")  # Update status text
        time_text.text(
            f"Total elapsed: {elapsed_time:.2f}s | Avg per PDF: {avg_time_per_pdf:.2f}s"
        )  # Update timing information

    # Display success message when all PDFs are processed
    st.success("Processing complete! Analysis reports have been saved.")
else:
    # Display a message prompting the user to provide input
    st.info("Enter a folder path and upload an Excel file to start processing.")
