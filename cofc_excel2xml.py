import sys
import os
import re
import argparse

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(script_dir)

from CoFCLib.COFCExcelCriticalDimensionWorksheetReader import COFCExcelCriticalDimensionWorksheetReader

# Default values specified at the top of the script
DEFAULT_EXCEL_FILE = r'C:\Users\magalang\Documents\examples\DNP\DEMO_DEVICE_XLS.xls'
DEFAULT_SELECTED_INFORMATION = 'Detail'
#DEFAULT_SELECTED_STRUCTURE_PATTERNS = 'THROUGH_PITCH,LINEARITY,TLL,LES,TJP'
DEFAULT_SELECTED_STRUCTURE_PATTERNS = ''

def main():
    parser = argparse.ArgumentParser(description="Process Excel file for COFC Critical Dimension Worksheet Reader")
    parser.add_argument('--excel_file', default=DEFAULT_EXCEL_FILE, help="Excel file path (default: %(default)s)")
    parser.add_argument('--selected_information', default=DEFAULT_SELECTED_INFORMATION, help="Selected information as comma-separated values (default: %(default)s)")
    parser.add_argument('--selected_structure_patterns', default=DEFAULT_SELECTED_STRUCTURE_PATTERNS, help="Selected structure patterns as command-separated values (default: %(default)s)")

    args = parser.parse_args()

    excel_file = args.excel_file
    selected_information = args.selected_information.split(',')
    selected_structure_patterns = args.selected_structure_patterns.split(',')

    # Create an instance of COFCExcelCriticalDimensionWorksheetReader
    cofc_cd = COFCExcelCriticalDimensionWorksheetReader(excel_file)
    cofc_cd.open_workbook()
    cofc_cd.selected_information = selected_information
    cofc_cd.selected_structure_patterns = selected_structure_patterns
    cofc_cd.create_category_tables()
    cofc_cd.generate_json_data()
    cofc_cd.cd_xml = cofc_cd.critical_dimension_json_to_xml('Critical_Dimension')

if __name__ == "__main__":
    main()