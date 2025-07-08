import os
import csv
import re
import pdfplumber

def process_pid(raw_pid):
    """
    Processes the raw PID string according to the specified rules.
    """
    # Extract all digits from the string
    digits = re.findall(r'\d+', raw_pid)
    if not digits:
        return None
    
    # Join all groups of digits
    full_number_str = "".join(digits)
    
    # Remove the last two digits
    processed_number_str = full_number_str[:-2]
    
    # Prepend "JB000"
    return f"JB000{processed_number_str}"

def extract_invoice_data(pdf_path):
    """
    Extracts all required data from a single PDF invoice using table data.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()

        # Extract data from the specific table cells
        invoice_number = tables[0][1][1]
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
    invoice_dir = r"C:\Users\Joey\Desktop\inv"
    output_csv = os.path.join(invoice_dir, "invoices_manual.csv")
    
    pdf_files = [f for f in os.listdir(invoice_dir) if f.lower().endswith('.pdf')]
    
    with open(output_csv, 'w', newline='') as csvfile:
        fieldnames = [
            "Project #", "PID # / Job ID #", "Invoice #", "Total Amount", "P.O. #",
            "Constant", "Code", "Quantity", "Rate", "Amount"
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for pdf_file in pdf_files:
            pdf_path = os.path.join(invoice_dir, pdf_file)
            print(f"Processing {pdf_file}...")
            try:
                data = extract_invoice_data(pdf_path)
                
                for item in data["Line Items"]:
                    row_data = {
                        "Project #": "473059",
                        "PID # / Job ID #": data["PID # / Job ID #"],
                        "Invoice #": data["Invoice #"],
                        "Total Amount": data["Total Amount"],
                        "P.O. #": data["P.O. #"],
                        "Constant": item["Constant"],
                        "Code": item["Code"],
                        "Quantity": item["Quantity"],
                        "Rate": item["Rate"],
                        "Amount": item["Amount"]
                    }
                    writer.writerow(row_data)
            except Exception as e:
                print(f"Could not process {pdf_file}: {e}")

    print(f"\nProcessing complete. Output saved to {output_csv}")

if __name__ == "__main__":
    main()