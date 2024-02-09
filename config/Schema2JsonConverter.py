#!/usr/bin/env python
# coding: utf-8

import json
import pandas as pd


class Schema2Json:

    def __init__(self, file_path,json_file):
        self.file_path: str = file_path
        self.json_file: str = json_file
        self.workbook: str = 'v3_cofcExtract_selected'
        self.current_sheet = None
        self.df = None
        self.open_workbook()

    def open_workbook(self):
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=self.workbook, header=None, usecols=None)
        except ValueError as e:
            print("e")
            exit(1)

    def generate_xpath_config(self):
        df = self.df
        start_row = df[df[0].str.startswith('Attribute Name in XML', na=False)].index[0]
        start_col = df.columns.get_loc(1)
        patterns = list(df.iloc[start_row, start_col-1:])
        patterns = [x for x in patterns if str(x) != 'nan']
        patterns = [s.replace(' ', '') for s in patterns]
        df.columns = patterns
        #patterns.remove('ReportedinTMAXML')
        patterns.remove('Mandatory?')
        patterns.remove('AttributeType')
        df = df[patterns]
        df = df[df['Selected'].isin([1,0])]
        df = df.iloc[1:].reset_index(drop=True)
        df['AttributeNameinXML'] = df['AttributeNameinXML'].str.replace(' ', '', regex=True)
        df['AttributeNameinXML'] = df['AttributeNameinXML'].replace('-', '/', regex=True)
        df['AttributeNameinXML'] = df['AttributeNameinXML'].replace('^', './/', regex=True)
        json_data = df.to_dict(orient='records')
        with open(self.json_file, 'w') as f:
            json.dump(json_data, f, indent=4)
        print(f"JSON file '{self.json_file}' created.")


is_schema2config = 1
if is_schema2config:
    ef: str = r'C:\Users\e14529\Documents\Carina\cofc_transformer\config\OS-006768-01_24 XML Schema_remake_20231016.xlsx'
    of: str = './OS-006768-01_24.xpaths.2023-10.json'
    cofc_s2c = Schema2Json(ef,of)
    json_data = cofc_s2c.generate_xpath_config()
