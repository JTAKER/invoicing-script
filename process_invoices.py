import os
import csv
import re
import pdfplumber
import datetime

def process_pid(raw_pid):
    """
    Processes the raw PID string according to the specified rules.
    """
    prefix = "JB"
    if raw_pid and "REQ" in raw_pid.upper():
        prefix = "REQ"

    # Extract all digits from the string
    digits = re.findall(r'\d+', raw_pid)
    if not digits:
        return None
    
    # Join all groups of digits
    full_number_str = "".join(digits)
    
    # If the number of digits is greater than 7, assume the last two are trailing and remove them.
    if len(full_number_str) > 7:
        core_number_str = full_number_str[:-2]
    else:
        core_number_str = full_number_str
        
    # Now, check the length of the core number and format accordingly.
    if len(core_number_str) == 6:
        return f"{prefix}0000{core_number_str}"
    else: # Handles 7-digit numbers and any other cases
        return f"{prefix}000{core_number_str}"

def extract_invoice_data(pdf_path):
    """
    Extracts all required data from a single PDF invoice using table data.
    """
    invoice_number = os.path.splitext(os.path.basename(pdf_path))[0]
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()

        # Extract data from the specific table cells
        main_table = tables[2]
        
        raw_pid = main_table[1][5]
        pid = process_pid(raw_pid)
        po_number = main_table[3][0]
        total_amount_str = main_table[-1][9]
        total_amount = total_amount_str.replace('$', '').replace(',', '')

        # Line items are in a single row, with values separated by newlines
        line_item_row = main_table[5]
        codes = line_item_row[0].split('\n')
        quantities = line_item_row[3].split('\n')
        rates = line_item_row[5].split('\n')
        amounts = line_item_row[10].split('\n')

        line_items = []
        for i in range(len(codes)):
            if codes[i]: # Ensure the row is not empty
                line_items.append({
                    "Constant": "1",
                    "Code": codes[i],
                    "Quantity": quantities[i],
                    "Rate": rates[i],
                    "Amount": amounts[i]
                })

        return {
            "PID # / Job ID #": pid,
            "Invoice #": invoice_number,
            "Total Amount": total_amount,
            "P.O. #": po_number,
            "Line Items": line_items
        }

def main():
    """
    Main function to process all PDF invoices in a directory and write to a CSV.
    """
    # Get batch letter
    while True:
        batch_letter = input("Which batch is this? (A-Z): ").upper()
        if len(batch_letter) == 1 and 'A' <= batch_letter <= 'Z':
            break
        print("Invalid input. Please enter a single letter from A to Z.")

    # Get invoice type
    while True:
        invoice_type_input = input("Is this an Expense or Commercial invoice? (e/c): ").lower()
        if invoice_type_input in ['e', 'c']:
            break
        print("Invalid input. Please enter 'e' for Expense or 'c' for Commercial.")

    invoice_type = "Expense" if invoice_type_input == 'e' else "Commercial"

    invoice_dir = os.getcwd()
    today_date = datetime.datetime.now().strftime("%Y-%m-%d")

    # Construct the filename
    filename = f"{today_date} Batch {batch_letter} {invoice_type}.csv"
    output_csv = os.path.join(invoice_dir, filename)
    
    pdf_files = [f for f in os.listdir(invoice_dir) if f.lower().endswith('.pdf')]
    
    with open(output_csv, 'w', newline='') as csvfile:
        # Write the custom first row
        custom_writer = csv.writer(csvfile)
        custom_writer.writerow(['InvoiceCSV_V2'])

        fieldnames = [
            "Vendor Number", "P2_JOB_ID", "Invoice Number", "Invoice Amount", "Purchase Order Number",
            "Purchase Order Number Line", "Labor Part Number", "Quantity Invoiced", "Unit Price", "Invoice Line Amount", "Notes"
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
        writer.writeheader()

        for pdf_file in pdf_files:
            pdf_path = os.path.join(invoice_dir, pdf_file)
            print(f"Processing {pdf_file}...")
            try:
                data = extract_invoice_data(pdf_path)
                
                for item in data["Line Items"]:
                    row_data = {
                        "Vendor Number": "473059",
                        "P2_JOB_ID": data["PID # / Job ID #"],
                        "Invoice Number": data["Invoice #"],
                        "Invoice Amount": data["Total Amount"],
                        "Purchase Order Number": data["P.O. #"],
                        "Purchase Order Number Line": item["Constant"],
                        "Labor Part Number": item["Code"],
                        "Quantity Invoiced": item["Quantity"],
                        "Unit Price": item["Rate"],
                        "Invoice Line Amount": item["Amount"],
                        "Notes": ""
                    }
                    writer.writerow(row_data)
            except Exception as e:
                print(f"Could not process {pdf_file}: {e}")

    print(f"\nProcessing complete. Output saved to {output_csv}")

if __name__ == "__main__":
    main()