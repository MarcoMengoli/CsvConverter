import csv
import os
import sys
import argparse
try:
    import xlsxwriter
except ImportError:
    print("xlsxwriter installation required. Run: pip install xlsxwriter")
    sys.exit(1)

def split_csv_to_excel(input_file, output_dir=None, rows_per_file=1000000):

    if not os.path.exists(input_file):
        print(f"Error: File {input_file} does not exist.")
        return
    
    if output_dir is None:
        output_dir = os.getcwd()
    else:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
    
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    
    print(f"Processing file {input_file}...")
    
    with open(input_file, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        
        headers = next(reader)
        
        current_file_index = 1
        current_row_count = 0
        current_workbook = None
        current_worksheet = None
        
        output_file = os.path.join(output_dir, f"{base_name}_{current_file_index}.xlsx")
        current_workbook = xlsxwriter.Workbook(output_file)
        current_worksheet = current_workbook.add_worksheet()
        
        for col_idx, header in enumerate(headers):
            current_worksheet.write(0, col_idx, header)
        
        for row in reader:
            current_row_count += 1
            
            for col_idx, cell_value in enumerate(row):
                current_worksheet.write(current_row_count, col_idx, cell_value)
            
            if current_row_count >= rows_per_file:
                current_workbook.close()
                
                print(f"Created file: {output_file}")
                
                current_file_index += 1
                current_row_count = 0
                output_file = os.path.join(output_dir, f"{base_name}_{current_file_index}.xlsx")
                current_workbook = xlsxwriter.Workbook(output_file)
                current_worksheet = current_workbook.add_worksheet()
                
                for col_idx, header in enumerate(headers):
                    current_worksheet.write(0, col_idx, header)
        
        current_workbook.close()
        print(f"Created file: {output_file}")
    
    print(f"Operation completed. {current_file_index} Excel files created in directory: {output_dir}")

def main():
    parser = argparse.ArgumentParser(description='Split a CSV file into multiple Excel files')
    parser.add_argument('input_file', help='Path to input CSV file')
    parser.add_argument('-o', '--output_dir', help='Directory for output files', default=None)
    parser.add_argument('-r', '--rows', type=int, help='Number of rows per Excel file', default=1000000)
    
    args = parser.parse_args()
    
    split_csv_to_excel(args.input_file, args.output_dir, args.rows)

if __name__ == "__main__":
    main()