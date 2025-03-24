import xml.etree.ElementTree as ET
import pandas as pd
import argparse
import os

def parse_excel_xml(xml_file, xlsx_file):
    """
    Parse an Excel XML file format (SpreadsheetML) and convert it to an Excel XLSX file.
    """
    try:
        # Parse XML
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        # Define namespaces
        namespaces = {
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
            'x': 'urn:schemas-microsoft-com:office:excel'
        }
        
        # Find all worksheets
        worksheets = {}
        
        # Extract data from all worksheets
        for worksheet in root.findall('.//{urn:schemas-microsoft-com:office:spreadsheet}Worksheet'):
            sheet_name = worksheet.get('{urn:schemas-microsoft-com:office:spreadsheet}Name', 'Sheet')
            
            # Find table data
            table = worksheet.find('.//{urn:schemas-microsoft-com:office:spreadsheet}Table')
            if table is None:
                continue
                
            rows_data = []
            max_cols = 0
            
            # Process each row
            for row in table.findall('.//{urn:schemas-microsoft-com:office:spreadsheet}Row'):
                row_data = []
                current_index = 0
                
                # Process each cell in the row
                for cell in row.findall('.//{urn:schemas-microsoft-com:office:spreadsheet}Cell'):
                    # Check if cell has an Index attribute (1-based)
                    cell_index = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                    if cell_index:
                        # Fill any gaps with empty values
                        index_val = int(cell_index) - 1  # Convert to 0-based
                        while current_index < index_val:
                            row_data.append("")
                            current_index += 1
                    
                    # Get cell data
                    data = cell.find('.//{urn:schemas-microsoft-com:office:spreadsheet}Data')
                    value = data.text if data is not None and data.text else ""
                    
                    # Handle data types
                    data_type = data.get('{urn:schemas-microsoft-com:office:spreadsheet}Type', 'String') if data is not None else 'String'
                    
                    # Convert types if needed
                    if data_type == 'Number':
                        try:
                            value = float(value) if value else 0.0
                        except (ValueError, TypeError):
                            value = 0.0
                    elif data_type == 'DateTime':
                        # Strip the time portion if it's just a date
                        if value and 'T00:00:00' in value:
                            value = value.split('T')[0]
                    
                    row_data.append(value)
                    current_index += 1
                
                # Only add non-empty rows
                if any(x for x in row_data if x):
                    rows_data.append(row_data)
                    max_cols = max(max_cols, len(row_data))
            
            # Ensure all rows have the same number of columns
            for i in range(len(rows_data)):
                if len(rows_data[i]) < max_cols:
                    rows_data[i].extend([""] * (max_cols - len(rows_data[i])))
            
            # If we have data, store it for this worksheet
            if rows_data and len(rows_data) > 1:
                # Convert to DataFrame - first row becomes header
                headers = rows_data[0]
                # Ensure header names are strings and unique
                headers = [str(h) if h else f"Column{i+1}" for i, h in enumerate(headers)]
                
                # Create DataFrame with the data rows
                df = pd.DataFrame(rows_data[1:], columns=headers)
                worksheets[sheet_name] = df
        
        # Write all worksheets to a single Excel file
        if worksheets:
            with pd.ExcelWriter(xlsx_file, engine='openpyxl') as writer:
                for sheet_name, df in worksheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"Successfully converted {xml_file} to {xlsx_file}")
            return True
        else:
            print(f"No data found in {xml_file}")
            return False
        
    except Exception as e:
        print(f"Error converting XML to Excel: {str(e)}")
        return False

def main():
    # Set up command-line argument parsing
    parser = argparse.ArgumentParser(description="Convert an Excel XML file to an XLSX file.")
    parser.add_argument("xml_file", help="Path to the input Excel XML file")
    parser.add_argument("xlsx_file", help="Path for the output XLSX file")
    args = parser.parse_args()

    # Validate input file exists
    if not os.path.isfile(args.xml_file):
        print(f"Error: Input file {args.xml_file} does not exist")
        return
    
    # Convert Excel XML to Excel XLSX
    parse_excel_xml(args.xml_file, args.xlsx_file)

if __name__ == "__main__":
    main()
