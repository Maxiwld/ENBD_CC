import os
import pandas as pd
import fitz  # PyMuPDF
import re
from datetime import datetime

def extract_transactions_from_pdf(pdf_path):
    transactions = []
    
    # Regular expression to match the full string of one transaction, including currency conversions and credited amounts
    transaction_pattern = re.compile(r"""
        (?P<full_tx>
            (?P<transaction_date>\d{2}/\d{2}/\d{4})\s+       # Transaction Date
            (?P<posting_date>\d{2}/\d{2}/\d{4})\s+          # Posting Date
            (?P<description>.*?(?:\s*\(.*?\))?)\s+          # Description, includes optional parentheses (e.g., currency conversion)
            (?P<amount>\d{1,3}(?:,\d{3})*\.\d{2}(?:CR)?)$   # Amount, supports 'CR' and commas
        )
    """, re.VERBOSE | re.MULTILINE)
    
    # Open and read the PDF using PyMuPDF
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            try:
                text = page.get_text()
                print(text)
                if text:
                    # Match entire rows of transactions
                    matches = transaction_pattern.finditer(text)
                    for match in matches:
                        full_tx_string = match.group("full_tx")
                        transaction_date = match.group("transaction_date")
                        posting_date = match.group("posting_date")
                        description = match.group("description")
                        amount = match.group("amount").replace(',', '')  # Remove commas for numeric conversion
                        is_credit = False
                        
                        if amount.endswith("CR"):
                            is_credit = True
                            amount = amount.replace("CR", "")  # Remove 'CR' for float conversion
                        
                        try:
                            amount = float(amount)
                            if is_credit:
                                amount = abs(amount)  # Keep credit amounts positive
                            else:
                                amount = -abs(amount)  # Convert debit amounts to negative
                        except ValueError:
                            continue  # Skip invalid rows
                        
                        transactions.append({
                            'Full TX String': full_tx_string,
                            'Transaction Date': transaction_date,
                            'Posting Date': posting_date,
                            'Description': description,
                            'Amount': amount,
                            'Type': 'Credit' if is_credit else 'Debit'
                        })
            except Exception as e:
                print(f"Error processing page: {e}")
                continue
    
    return transactions

def extract_metadata_from_pdf(pdf_path):
    metadata = {}
    
    # Open and read the PDF using PyMuPDF
    with fitz.open(pdf_path) as pdf:
        all_lines = []
        for page in pdf:
            try:
                text = page.get_text()
                lines = text.split("\n")  # Split text into lines
                all_lines.extend(lines)
            except Exception as e:
                print(f"Error processing page: {e}")
                continue
        print(all_lines)
      
        # Regex to match patterns
        card_number_pattern = re.compile(r"\d{4}\sXXXX\sXXXX\s\d{4}")
        date_pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
        number_pattern = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")
        
        for i, line in enumerate(all_lines):
            if card_number_pattern.match(line.strip()):  # Card number is anchor
                metadata["Card Number"] = line.strip()
                
                # Extract statement period
                if i + 1 < len(all_lines):
                    metadata["Statement Period"] = all_lines[i + 1].strip()
                    period = all_lines[i + 1].strip()
                    if "to" in period:
                        start_date, end_date = period.split(" to ")
                        metadata["Statement Start Date"] = start_date
                        metadata["Statement End Date"] = end_date

                # Extract credit limit and available credit limit
                if i + 2 < len(all_lines) and number_pattern.match(all_lines[i + 2].strip()):
                    metadata["Credit Limit"] = all_lines[i + 2].strip()
                if i + 3 < len(all_lines) and number_pattern.match(all_lines[i + 3].strip()):
                    metadata["Available Credit Limit (AED)"] = all_lines[i + 3].strip().replace(",", "")

                # Search for the first two dates after available credit limit
                dates = []
                for j in range(i + 4, len(all_lines)):
                    if date_pattern.match(all_lines[j].strip()):
                        dates.append(all_lines[j].strip())
                    if len(dates) == 2:  # Only first two dates needed
                        break
                if len(dates) == 2:
                    metadata["Statement Date"] = dates[0]
                    metadata["Payment Due Date"] = dates[1]
                
                # Extract minimum payment due
                if j + 1 < len(all_lines) and number_pattern.match(all_lines[j + 1].strip()):
                    metadata["Minimum Payment Due"] = all_lines[j + 1].strip().replace(",", "")

                # Extract statement summary fields using "Closing Balance" anchor
                closing_balance_anchor = "Closing Balance"
                if closing_balance_anchor in all_lines:
                    index = all_lines.index(closing_balance_anchor)
                    if index + 1 < len(all_lines):
                        metadata["Previous Statement Due"] = all_lines[index + 1].strip().replace(",", "")
                    if index + 2 < len(all_lines):
                        metadata["Purchase / Cash Advance"] = all_lines[index + 2].strip().replace(",", "")
                    if index + 3 < len(all_lines):
                        metadata["Interest / Other Charges"] = all_lines[index + 3].strip().replace(",", "")
                    if index + 4 < len(all_lines):
                        metadata["Payments / Credits"] = all_lines[index + 4].strip().replace(",", "")
                    if index + 5 < len(all_lines):
                        metadata["Current Balance"] = all_lines[index + 5].strip().replace(",", "")

                break  # Stop after first card number match
    
    return metadata

def convert_metadata_dates(metadata):
    def convert_date_format(date_str):
        try:
            return datetime.strptime(date_str, "%d-%b-%y").strftime("%d/%m/%Y")
        except ValueError:
            return date_str  # Return the original if conversion fails

    if "Statement Start Date" in metadata:
        metadata["Statement Start Date"] = convert_date_format(metadata["Statement Start Date"])
    if "Statement End Date" in metadata:
        metadata["Statement End Date"] = convert_date_format(metadata["Statement End Date"])

    return metadata

def process_pdfs_in_folder(folder_path, output_excel):
    all_metadata = []
    writer = pd.ExcelWriter(output_excel, engine="openpyxl")

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    print(f"Found {len(pdf_files)} PDFs in the folder")

    for idx, file_name in enumerate(pdf_files):
        pdf_path = os.path.join(folder_path, file_name)
        print(f"Parsing file {idx + 1}/{len(pdf_files)} ({(idx + 1) / len(pdf_files) * 100:.2f}%) - {file_name}")

        metadata = extract_metadata_from_pdf(pdf_path)
        metadata = convert_metadata_dates(metadata)
        metadata["File Name"] = file_name
        all_metadata.append(metadata)

        # Extract transactions and write to a new sheet
        transactions = extract_transactions_from_pdf(pdf_path)
        transactions_df = pd.DataFrame(transactions)
        sheet_name = f"Statements_{file_name.split('_')[2]}"  # Shorten sheet name if necessary
        transactions_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Consolidate metadata into a single sheet
    metadata_df = pd.DataFrame(all_metadata)
    metadata_df.to_excel(writer, sheet_name="Metadata", index=False)

    # Save the Excel file
    writer.close()

# Example usage
folder_path = r"C:\Users\User\YourPath\ENBD\CC Statements"                   #Adjust to your folder path
output_excel = r"C:\Users\User\YourPath\ENBD\Consolidated_Statements.xlsx"   #Adjust to your folder path
process_pdfs_in_folder(folder_path, output_excel)
