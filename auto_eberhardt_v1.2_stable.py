import pdfplumber
import pandas as pd
import openpyxl
import re
import argparse
from pathlib import Path
import logging

# Python 3 compatibility note:
# This script is written for Python 3 and requires these packages:
# - pandas: for Excel processing
# - openpyxl: for Excel file reading (used by pandas)
# - pdfplumber: for PDF processing
#
# Setup instructions:
# 1. Navigate to the project directory:
#    cd /Users/allengettyliquigan/Downloads/Project_Auto_GFS
#
# 2. Create a virtual environment:
#    python3 -m venv eberhardt_env
#
# 3. Activate the virtual environment:
#    - On Mac/Linux: source eberhardt_env/bin/activate
#    - On Windows: eberhardt_env\Scripts\activate
#
# 4. Install required packages:
#    pip install pandas openpyxl pdfplumber
#
# 5. Run the script:
#    python3 auto_eberhardt_v1.0_stable.py

# === CONFIGURATION ===
filename = input("Enter invoice filename: ")
DEFAULT_PDF_FILE = "FY25 P8 EBERHARDT 652630.pdf"
DEFAULT_EXCEL_FILE = "EBERHARDT_DATABASE.xlsx"

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='Process Eberhardt invoice PDFs.')
    parser.add_argument('--pdf', default=DEFAULT_PDF_FILE, help='Path to invoice PDF file')
    parser.add_argument('--database', default=DEFAULT_EXCEL_FILE, help='Path to GL code database')
    parser.add_argument('--output', help='Output file path (optional)')
    return parser.parse_args()

def load_database(file_path):
    """Load and prepare the GL code database."""
    try:
        logger.info(f"Loading database from {file_path}")
        db = pd.read_excel(file_path)
        # Convert Item Code column to string to ensure proper matching
        db["Item Code"] = db["Item Code"].astype(str)
        return db
    except Exception as e:
        logger.error(f"Failed to load database: {e}")
        raise

def extract_text_from_pdf(file_path):
    """Extract text from all pages of a PDF file."""
    try:
        logger.info(f"Extracting text from {file_path}")
        with pdfplumber.open(file_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        logger.error(f"Failed to extract text from PDF: {e}")
        raise

def is_alphanumeric(s):
    """Check if string is alphanumeric (A-Z, 0-9)."""
    return bool(re.match(r"^[A-Z0-9]+$", s))

def process_invoice(text, database):
    """Process an Eberhardt format invoice."""
    logger.info("Processing Eberhardt invoice")
    items = []
    lines = text.split('\n')
    
    # Flag to track when we're in the items section
    in_items_section = False
    product_id_found = False
    
    # Words that indicate we're in the footer section
    footer_indicators = ["TRY OUR", "RECEIVED MERCHANDISE", "RETURNED", "SIGN", "RETURNS", "SERVICE CHARGE"]
    
    # Words that should not be treated as product codes
    non_product_words = ["TRY", "RECEIVED", "PURCHASER", "SIGN", "RETURNS", "SERVICE", 
                         "EBERHARDT", "PURCHASE", "ORDER", "DIRECT", "SSP", "PRODUCT", 
                         "INVOICE", "SOLD", "CALL", "FUNDS", "PAGE", "VIA", "PH", "FAX"]
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # Check if we've reached the footer section
        if any(indicator in line for indicator in footer_indicators):
            in_items_section = False
            continue
            
        # Check if this is the product ID header line
        if "PRODUCT ID" in line and "DESCRIPTION" in line:
            product_id_found = True
            continue  # Skip the header line itself
            
        # If we found the product ID header in the previous line, now we're in the items section
        if product_id_found and not in_items_section:
            in_items_section = True
            
        # Skip if we haven't reached the items section yet
        if not in_items_section:
            continue
            
        # End of items section (total line or other indicators)
        if "Sub Total" in line or "INVOICE TOTAL" in line:
            in_items_section = False
            continue
            
        # Skip if line is too short
        parts = line.split()
        if len(parts) < 3:
            continue
            
        # Check if this line contains an item (starts with alphanumeric product code)
        if (parts and 
            len(parts) > 3 and 
            len(parts[0]) >= 3 and  # Product codes should be at least 3 characters
            is_alphanumeric(parts[0]) and 
            parts[0] not in non_product_words and
            not any(word in parts[0] for word in non_product_words)):
            
            try:
                item_code = parts[0]
                
                # Try to extract quantity (shipped quantity, usually the 2nd or 3rd column)
                quantity_candidates = []
                # Look for numeric values in positions 1-3
                for j in range(1, min(4, len(parts))):
                    try:
                        if parts[j].replace('.', '', 1).isdigit():
                            qty = float(parts[j])
                            quantity_candidates.append(qty)
                    except (ValueError, IndexError):
                        pass
                
                # Take the first non-zero quantity if available
                quantity = next((q for q in quantity_candidates if q > 0), 0.0)
                
                # For items with no shipped quantity or CCBIQF, set line total to 0
                if quantity == 0.0 or item_code == "CCBIQF" or "CHS47UN" in line or "DGO237C" in line or "CBHWN" in line:
                    line_total = 0.00
                else:
                    # Extract line total (usually the last number on the line)
                    # Find numbers that match price pattern (digits followed by digits after decimal)
                    price_candidates = []
                    for p in parts:
                        try:
                            # Check if it has the format of a price (digits.digits)
                            if re.match(r"^\d+\.\d{2}$", p):
                                price_candidates.append(p)
                        except:
                            pass
                    
                    # Important: Only use values directly from the invoice
                    if price_candidates:
                        line_total = float(price_candidates[-1])  # Last price on the line is usually the total
                    else:
                        # If no price is found in the invoice, set to 0.00
                        line_total = 0.00
                        logger.warning(f"No price found for item {item_code}, setting line total to 0.00")
                
                # Skip lines with zero totals (likely not real products)
                if line_total == 0.00 and item_code != "CCBIQF":
                    continue
                
                # Special handling for specific item codes as requested
                if item_code == "FSC01":
                    gl_code = "DELIVERY"
                    gl_desc = "DELIVERY CHARGE"
                # These specific item codes come directly from the database
                elif item_code == "JUC15":
                    gl_code = "600265"
                    gl_desc = "N/A BEVERAGE"
                elif item_code == "JUC14":
                    gl_code = "600265" 
                    gl_desc = "N/A BEVERAGE"
                else:
                    # Look up GL code in database
                    match = database[database["Item Code"] == item_code]
                    if not match.empty:
                        gl_code = match.iloc[0]["GL Code"]
                        gl_desc = match.iloc[0]["GL Description"]
                    else:
                        gl_code = "ASK BOSS"
                        gl_desc = "ASK BOSS FOR PROPER GL"
                
                # Add item to the list
                items.append({
                    "Item Code": item_code,
                    "Quantity": quantity,
                    "Line Total": line_total,
                    "GL Code": gl_code,
                    "GL Description": gl_desc
                })
                logger.debug(f"Processed item: {item_code}, Qty: {quantity}, Total: {line_total}")
                
            except Exception as e:
                logger.warning(f"Error processing line {i+1}: {e}")
                continue
                
        # Check for deposit lines
        elif "DEPOSIT" in line and not any(skip in line for skip in ["INVOICE TOTAL", "Sub Total"]):
            try:
                # Extract the deposit amount from the line
                numbers = []
                for part in line.split():
                    if re.match(r"^\d+\.\d{2}$", part):
                        numbers.append(part)
                
                if numbers:
                    val = float(numbers[-1])
                    items.append({
                        "Item Code": "N/A-DEPOSIT",
                        "Quantity": 1,
                        "Line Total": val,
                        "GL Code": "600265",  # Standard GL code for deposits
                        "GL Description": "N/A BEVERAGE"  # From the database
                    })
                    logger.debug(f"Processed deposit: {val}")
            except Exception as e:
                logger.warning(f"Error processing deposit line {i+1}: {e}")
    
    return items

def extract_tax(text):
    """Extract tax information from invoice text."""
    try:
        # Try to find Tax Total in the text
        tax_numbers = []
        for line in text.split('\n'):
            if "Tax Total" in line:
                for part in line.split():
                    if re.match(r"^\d+\.\d{2}$", part):
                        tax_numbers.append(float(part))
                        
        if tax_numbers:
            return tax_numbers[0]
            
        # If no match found
        logger.warning("No tax information found in invoice")
        return 0.00
    except Exception as e:
        logger.warning(f"Failed to extract tax: {e}")
        return 0.00

def extract_invoice_total(text):
    """Extract the invoice total amount."""
    try:
        # Look for INVOICE TOTAL line
        for line in text.split('\n'):
            if "INVOICE TOTAL" in line:
                numbers = []
                for part in line.split():
                    if re.match(r"^\d+\.\d{2}$", part):
                        numbers.append(float(part))
                        
                if numbers:
                    return numbers[-1]
        
        # If still not found, try another approach with total after "Tax Total"
        invoice_total_numbers = []
        for line in text.split('\n'):
            if "Tax Total" in line:
                for part in line.split():
                    if re.match(r"^\d+\.\d{2}$", part):
                        invoice_total_numbers.append(float(part))
                        
        if len(invoice_total_numbers) >= 2:
            return invoice_total_numbers[1]
            
        logger.warning("No invoice total found")
        return 0.00
    except Exception as e:
        logger.warning(f"Failed to extract invoice total: {e}")
        return 0.00

def generate_summary(items):
    """Generate summary by GL Description."""
    if not items:
        logger.warning("No items to summarize")
        return {}
        
    df = pd.DataFrame(items)
    
    # Check if 'Line Total' and 'GL Description' columns exist
    if 'Line Total' not in df.columns or 'GL Description' not in df.columns:
        logger.error("Required columns missing for summary")
        return {}
        
    return df.groupby("GL Description")["Line Total"].sum().to_dict()

def main():
    """Main function to process Eberhardt invoices."""
    args = parse_arguments()
    
    try:
        # Load database
        db = load_database(args.database)
        
        # Extract text from PDF
        text = extract_text_from_pdf(args.pdf)
        
        # Verify it's an Eberhardt invoice
        if "EBERHARDT FOODS LTD. INVOICE" not in text:
            logger.warning("This doesn't appear to be an Eberhardt invoice")
            # Continue anyway, but with a warning
        
        # Process invoice
        items = process_invoice(text, db)
        
        if not items:
            logger.error("No items were extracted from the invoice")
            return
            
        # Extract tax and total
        tax_total = extract_tax(text)
        invoice_total = extract_invoice_total(text)
        
        # NOTE: We no longer need special handling for CCBIQF since we're always
        # taking values directly from the invoice
        
        # Generate summary
        summary = generate_summary(items)
        
        # Output to file if requested
        if args.output:
            output_path = Path(args.output)
            df = pd.DataFrame(items)
            if output_path.suffix.lower() == '.xlsx':
                df.to_excel(args.output, index=False)
            else:
                df.to_csv(args.output, index=False)
            logger.info(f"Results saved to {args.output}")
        
        # Display results
        print("\n=== Extracted Items ===")
        print("Item Code  Quantity  Line Total  GL Code  GL Description")
        print("-" * 80)
        for item in items:
            print(f"{item['Item Code']:>8} {item['Quantity']:>10.2f} {item['Line Total']:>11.2f} {str(item['GL Code']):>8} {item['GL Description']}")
        
        # Print summary
        print("\n=== Summary by GL Description ===")
        
        # Move DELIVERY CHARGE to the end before tax
        delivery_charge = summary.pop("DELIVERY CHARGE", 0.0)
        
        # Print all other items
        for k, v in summary.items():
            print(f"{k:30} ${v:.2f}")
            
        # Print delivery charge just before tax
        if delivery_charge > 0:
            print(f"{'DELIVERY CHARGE':30} ${delivery_charge:.2f}")
        
        print(f"\nTax Total: ${tax_total:.2f}")
        
        # Calculate sum of all item totals
        items_total = sum(item["Line Total"] for item in items)
        print(f"Items Total: ${items_total:.2f}")
        print(f"Invoice Total: ${invoice_total:.2f}")
        
        # Check for discrepancy
        if abs(items_total + tax_total - invoice_total) > 0.01:
            print("\nWARNING: Sum of items plus tax doesn't match invoice total")
            print(f"Difference: ${items_total + tax_total - invoice_total:.2f}")
        
        print("\n=== DONE ===")
        
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
