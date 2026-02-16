"""
 * Author: Alejandro Ramírez
 * Project: excel-to-csv-converter
 * Logic: Professional CLI engine for high-fidelity conversion of Excel 
 * workbooks to CSV, featuring batch sheet processing and automated path resolution.
 """

import os
import sys
import argparse
import pandas as pd
from typing import List, Optional

class ExcelConverterLogic:
    def __init__(self, verbose: bool = False):
        self.verbose = verbose

    def log(self, message: str):
        if self.verbose:
            print(f"[LOG] {message}")

    def convert(self, input_path: str, output_path: Optional[str] = None, sheet_name: Optional[str] = None, all_sheets: bool = False):
        """
        Orchestrates the conversion process.
        """
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Source system failure: File not found at {input_path}")

        self.log(f"Initializing extraction for: {input_path}")
        
        try:
            excel_file = pd.ExcelFile(input_path)
            sheets = excel_file.sheet_names
            
            if all_sheets:
                self._convert_all_sheets(excel_file, sheets, input_path)
            else:
                target_sheet = sheet_name if sheet_name else sheets[0]
                self._convert_single_sheet(excel_file, target_sheet, output_path, input_path)
                
        except Exception as e:
            print(f"[ERROR] Extraction failed: {str(e)}")
            sys.exit(1)

    def _convert_single_sheet(self, excel_file, sheet_name, output_path, input_path):
        if not output_path:
            base_name = os.path.splitext(input_path)[0]
            output_path = f"{base_name}_{sheet_name}.csv" if sheet_name != excel_file.sheet_names[0] else f"{base_name}.csv"

        self.log(f"Processing sheet: '{sheet_name}' -> {output_path}")
        df = excel_file.parse(sheet_name)
        df.to_csv(output_path, index=False)
        print(f"[SUCCESS] Exported: {output_path}")

    def _convert_all_sheets(self, excel_file, sheets, input_path):
        self.log(f"Batch mode activated. Processing {len(sheets)} sheets.")
        base_name = os.path.splitext(input_path)[0]
        
        for sheet in sheets:
            out = f"{base_name}_{sheet}.csv"
            self._convert_single_sheet(excel_file, sheet, out, input_path)

def main():
    parser = argparse.ArgumentParser(description="Alejandro Ramírez | Excel-to-CSV High-Performance Converter")
    parser.add_argument("input", help="Path to the source Excel file (.xlsx, .xls)")
    parser.add_argument("-o", "--output", help="Destination path for the CSV (default: same as input)")
    parser.add_argument("-s", "--sheet", help="Specific sheet name to convert (default: first sheet)")
    parser.add_argument("-a", "--all", action="store_true", help="Convert all sheets in the workbook")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable detailed logging")

    args = parser.parse_args()

    converter = ExcelConverterLogic(verbose=args.verbose)
    converter.convert(args.input, args.output, args.sheet, args.all)

if __name__ == "__main__":
    main()
