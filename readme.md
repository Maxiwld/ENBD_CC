# Credit Card PDF Parser

## Overview
This script processes credit card statements from Emirates NBD in PDF format (as of 20.12.2024), extracting both metadata and transaction history. The output is consolidated into a single Excel file, where each PDF's metadata is summarized in a metadata sheet, and its transaction history is written to a dedicated sheet.

## Features
1. **Metadata Parsing**
   - Extracts information such as card number, statement period, credit limits, and payment details.
   - Formats dates consistently (e.g., `11/10/2024`).

2. **Transaction Parsing**
   - Reads individual transactions, including details like dates, descriptions, amounts, and types (credit/debit).
   - Supports multiple pages and inconsistent formatting within PDFs.

3. **Excel Consolidation**
   - Metadata for all PDFs is stored in a single sheet named `Metadata`.
   - Transaction history is stored in separate sheets named using the date within the file name (e.g., `Statements_10012024`).

4. **Progress Logging**
   - Displays progress, including the number of PDFs found and parsing status.

## Setup

### Dependencies
Ensure the following Python libraries are installed:
- `os`: For folder and file handling.
- `pandas`: For handling and writing tabular data.
- `fitz` (PyMuPDF): For extracting text from PDFs.
- `re`: For regular expression matching.
- `datetime`: For date manipulation.

### Installation
Run the following to install dependencies:
```bash
pip install pandas pymupdf openpyxl
```

### Folder Structure
Place all the PDF files to be processed in the target folder (e.g., `CC Statements`).

## Usage
1. Update the `folder_path` variable with the path to your target folder containing PDFs.
2. Update the `output_excel` variable with the desired path for the output Excel file.
3. Run the script.

## Output
1. **Excel File**
   - A single Excel file containing:
     - `Metadata` sheet summarizing information from all PDFs.
     - Separate sheets for each file's transactions, named based on the date extracted from the file name.

## Code Flow
1. **Folder Scan**
   - Scans the specified folder for PDFs.
   - Logs the number of files found.

2. **Metadata Parsing**
   - Extracts metadata from each PDF, including:
     - Card number.
     - Statement period and derived dates.
     - Credit limits.
     - Payment information.
     - Statement summary (e.g., balances, charges).

3. **Transaction Parsing**
   - Extracts and processes each PDF's transaction history.
   - Writes transactions to a separate Excel sheet for each file.

4. **Excel Generation**
   - Metadata and transaction sheets are consolidated into a single Excel file.

## Notes
- Ensure no files are open in the target folder during execution.
- Long file names are shortened for Excel sheet naming to avoid errors.
- The script automatically handles malformed PDFs by skipping over them with appropriate error messages.

