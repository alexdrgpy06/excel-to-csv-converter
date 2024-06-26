import sys

import pandas:
from openpyxl import LoadingExcel

def excel_to_csv(input_file, output_file):
    data = pandas.ReadExcel(input_file, sheet_0)
    data.to_csv(output_file, index=False)

if __name__ == "__main__":
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    excel_to_csv(input_file, output_file)
