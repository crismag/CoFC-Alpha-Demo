import sys
import os
import argparse

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(script_dir)

from CoFCLib.COFCExcelCopyProtect import COFCCopyProtect

# Default values specified at the top of the script
use_pandas_only = False
#DEFAULT_EXCEL_INFILE = r'C:\Users\magalang\Documents\COA_DEMO.protected.xls'
DEFAULT_EXCEL_INFILE = ''

if use_pandas_only:
    DEFAULT_EXCEL_OUTFILE = 'copied_excel_worksheet.xlsx'
else:
    DEFAULT_EXCEL_OUTFILE = 'copied_excel_worksheet.xls'

DEFAULT_SHEETS_OF_INTEREST = ['Statistics']


def main():
    parser = argparse.ArgumentParser(description="Copy protected Excel worksheets with optional settings")

    parser.add_argument('excel_infile', help="Input Excel file path")
    parser.add_argument('--excel_outfile', default=DEFAULT_EXCEL_OUTFILE, help="Output Excel file.")
    parser.add_argument('--sheets_of_interest', default=None, help="Sheets of interest as comma-separated values")

    args = parser.parse_args()

    excel_infile = args.excel_infile
    excel_outfile = args.excel_outfile
    sheets_of_interest = args.sheets_of_interest


    # Set default values if not provided
    if excel_outfile is None:
        base_filename, file_extension = os.path.splitext(os.path.basename(excel_infile))
        excel_outfile = f'{base_filename}_copied{file_extension}'

    if sheets_of_interest is not None:
        sheets_of_interest = sheets_of_interest.split(',')

    copier = COFCCopyProtect(excel_infile, excel_outfile, sheets_of_interest, use_pandas_only)
    copier.copy_worksheets()
    worksheets = copier.get_worksheets()
    print("Worksheets:", worksheets)


if __name__ == "__main__":
    main()