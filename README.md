# Excel Data Extraction and Report Generation

## Project Description

This project automates the extraction of data from multiple Excel invoices and generates PDF reports for each invoice. The process involves reading Excel files, formatting the data into a table layout, calculating totals, and creating PDF files with the extracted information.

## Process

1. **File Discovery:**
   - Use the `glob` module to locate all `.xlsx` files in the `Invoices` directory.

2. **Data Extraction:**
   - For each Excel file, read data from the "Sheet 1" using Pandas.

3. **PDF Creation:**
   - Extract invoice number and date from the filename.
   - Initialize a new PDF document with the FPDF library.
   - Add invoice number and date as the header in the PDF.
   - Create a table header and populate it with data from the Excel file.
   - Add rows to the table for each entry in the Excel file.
   - Calculate and display the total price of the invoice.
   - Include the company name and logo in the PDF.

4. **Output:**
   - Save each generated PDF file to the `PDFs` directory with a filename corresponding to the original Excel file.

## Technology Used

- **Pandas:** Used for reading data from Excel files.
- **FPDF:** Used for generating and formatting PDF documents.
- **Python:** The programming language used for scripting the extraction and report generation.

## What I Learned

- **Data Handling:**
  - Gained experience in reading and processing data from Excel files using Pandas.
  - Improved skills in data manipulation and aggregation.

- **PDF Generation:**
  - Learned how to use the FPDF library to create and format PDF documents programmatically.
  - Acquired knowledge in designing and structuring PDF reports with dynamic content.

- **File Management:**
  - Implemented file path handling and automated PDF generation based on file names.
  - Enhanced understanding of file I/O operations in Python.

## Future Insights

- **Enhancements:**
  - Add support for additional Excel sheet names or formats.
  - Implement more complex report layouts and styles.
  - Allow user-defined templates for generating PDF reports.

- **Error Handling:**
  - Improve error handling for scenarios where files are missing, have incorrect formats, or contain invalid data.
  - Add validation and user feedback for input files and processing stages.

- **Integration:**
  - Explore integrating with other data sources or formats, such as CSV or databases.
  - Consider adding a GUI or web interface for easier file selection and report generation.

Feel free to explore the code to understand the PDF generation process:
- `main.py` for the core functionality of data extraction and PDF creation.

Happy reporting!
