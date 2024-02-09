#!/usr/bin/env python
# coding: utf-8
# ------------------------------------------------
# Cristopher Magalang, 2023
# Parsing CoFC Critical Dimension page.
# ------------------------------------------------

import xml.etree.ElementTree as ET
from math import isnan
from typing import Dict, Any
import re
import pandas as pd
import xml.dom.minidom as minidom
import json

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)


class COFCExcelCriticalDimensionWorksheetReader:
    main_row: Dict[Any, Any]

    def __init__(self, file_path):
        self.file_path: str = file_path
        self.workbook: str = 'Critical Dimension'
        self.selected_information = ['Detail', 'Result', 'Specification']
        self.selected_structure_patterns = ['default']
        self.current_sheet = None
        self.main_row = {}
        self.df = None
        self.table = None
        self.df_filtered = None
        self.combined_category_table = None
        self.patterns = None
        self.json_data = None

    def open_workbook(self, workbook="Critical Dimension"):
        """
        :type workbook:str
        """
        self.workbook = workbook
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=self.workbook, header=None, usecols=None,
                                    keep_default_na=False, na_values=['','nan','NaN'])
        except ValueError as e:
            print("e")
        except Exception as e:
            print("e")
        try:
            self.df.replace('\n', ' ', regex=True, inplace = True)
        except Exception as e:
            pass


    @property
    def combine_details_row(self):
        table = self.table
        patterns = self.patterns
        # Initialize the new row dictionary with default values
        new_row = {"Categories": "Detail", "SubCategory": "DetailsList"}

        try:
            # Filter the table to get only rows with "Categories" == "Detail"
            detail_rows = table[table['Categories'] == 'Detail']

            # Loop through the given patterns and extract the values for each
            for pat in patterns:
                # Use pandas Series.dropna() to remove any NaN values in the column
                foo_values = detail_rows[pat].dropna().values.tolist()

                # Update the new row dictionary with the extracted values for the pattern
                new_row.update({pat: foo_values})
        except Exception as e:
            # Handle any exceptions that may occur during the function execution
            print(f"Error occurred while combining detail rows: {e}")
            new_row = {}  # Return an empty dictionary if an error occurs

        return new_row

    def combine_results_rows(self):
        table = self.table
        pattern = 'Result'
        df = table[table['Categories'] == pattern]
        result_row = table[table['Categories'] == pattern]
        new_row = {"Categories": "Result", "SubCategory": "Information_Results"}
        for col in df.columns[2:]:
            subcats = df['SubCategory'].unique()
            new_json = {}
            for subcat in subcats:
                val = df.loc[df['SubCategory'] == subcat, col].iloc[0]
                new_json.update({subcat: val})
            new_row.update({col: json.dumps(new_json)})
        return new_row

    def combine_specifications_rows(self):
        table = self.table
        pattern = 'Specification'
        df = table[table['Categories'] == pattern]
        result_row = table[table['Categories'] == pattern]
        new_row = {"Categories": "Specification", "SubCategory": "Information_Specification"}
        for col in df.columns[2:]:
            subcats = df['SubCategory'].unique()
            new_json = {}
            for subcat in subcats:
                val = df.loc[df['SubCategory'] == subcat, col].iloc[0]
                new_json.update({subcat: val})
            new_row.update({col: json.dumps(new_json)})
        return new_row

    def create_category_tables(self):
        df = self.df
        # Get the patterns as column names
        for keys in ['Device', 'Layer', 'Tool']:
            row = df[df[1] == keys].index[0]
            self.main_row.update({keys: str(df.loc[row][2])})

        start_row = df[df[1].str.startswith('Information', na=False)].index[0]
        start_col = df.columns.get_loc(1)
        patterns = []
        patterns = list(df.iloc[start_row, start_col + 1:])
        patterns = [x for x in patterns if str(x) != 'nan']
        if patterns[0] == 'No.':
            patterns.pop(0)
        if patterns[0] == 'Pattern':
            patterns.pop(0)

        self.patterns = None
        self.patterns = patterns

        item_count = {}
        # Count the occurrences of each pattern item in the list
        for item in patterns:
            if item in item_count:
                item_count[item] += 1
            else:
                item_count[item] = 1

        # Append suffix to duplicates
        for i in range(len(patterns)):
            item = patterns[i]
            if item_count[item] > 1:
                suffix = '_#DUP_' + str(item_count[item] - 1)
                patterns[i] = str(item) + suffix
                item_count[item] += 1

        # Define column names
        column_names = ['Categories', 'SubCategory'] + patterns

        # Slice the table from the start row and column
        table = df.iloc[start_row - 1:, start_col:len(column_names) + 1]

        # Drop the first row of the table
        table = table.iloc[1:].reset_index(drop=True)

        # Rename columns
        table.columns = column_names

        # Replace NaN values in Categories with the category name
        table['Categories'] = table['Categories'].fillna(method='ffill')

        # Reset the index to ensure all values are unique
        table = table.reset_index(drop=True)
        self.table = table
        new_table = table

        # create a new dataframe with rows where Categories does not equal Detail, Result, or Specification
        for cat in ['Detail','Result','Specification']:
            new_table = new_table[new_table['Categories'] != cat].copy()

        # combine the "Detail" rows
        if 'Detail' in self.selected_information:
            try:
                new_detail_row = self.combine_details_row
                new_table = pd.concat([new_table, pd.DataFrame([new_detail_row])], ignore_index=True)
            except Exception as e:
                print(f"Detail Row Error: {str(e)}")

        # combine the "Result" rows
        if 'Result' in self.selected_information:
            try:
                new_result_row = self.combine_results_rows()
                new_table = pd.concat([new_table, pd.DataFrame([new_result_row])], ignore_index=True)
            except Exception as e:
                print(f"Result Row Error: {str(e)}")

        # combine the "Specifications" rows
        if 'Specifications' in self.selected_information:
            try:
                new_spec_row = self.combine_specifications_rows()
                new_table = pd.concat([new_table, pd.DataFrame([new_spec_row])], ignore_index=True)
            except Exception as e:
                print(f"Specifications Row Error: {str(e)}")

        self.combined_category_table = new_table
        return new_table

    def critical_dimension_json_to_xml(self, root_name='Critical_Dimension'):
        json_data = self.json_data
        root = ET.Element(root_name)

        for structure_name, details in json_data.items():
            site_element = ET.SubElement(root, 'CD_Site')

            name_element = ET.SubElement(site_element, 'Name')
            name_element.text = structure_name
            for key, value in details.items():
                if key == 'Pattern':
                    continue
                if key == 'DetailsList':
                    details_element = ET.SubElement(site_element, 'Details')
                    for value in details['DetailsList']:
                        value_element=ET.SubElement(details_element,'value')
                        value_element.text = str(value)
                elif key.startswith('Information_') or key.startswith('Information_'):
                    key = key.replace('Information_', '')
                    info_element = ET.SubElement(site_element, key)
                    for k, v in value.items():
                        k = k.replace(' ', '_')
                        if k == '3_sigma' or k == '3sigma':
                            k = 'Stat_3_sigma'
                        info_element.set(k, str(v))
                else:
                    sub_element = ET.SubElement(site_element, key)
                    sub_element.text = str(value)

        xml_string = ET.tostring(root, encoding='unicode')
        xml_dom = minidom.parseString(xml_string)
        pretty_xml = xml_dom.toprettyxml(indent='  ')
        return pretty_xml

    def generate_json_data(self):
        df = self.combined_category_table

        # Transpose the DataFrame
        dft = df.transpose().reset_index(drop=True)

        # Set the column names to the values in the second row
        column_names = list(dft.iloc[1])
        dft.columns = column_names

        # Remove the second row
        dft = dft.iloc[2:]

        # Reset the index
        dft = dft.reset_index(drop=True)

        if 'default' in self.selected_structure_patterns or self.selected_structure_patterns == []:
            self.df_filtered = dft
        else:
            self.df_filtered = dft[dft['Pattern'].str.contains('|'.join(self.selected_structure_patterns), regex=True)]
            self.df_filtered.reset_index(drop=True)

        # Convert the DataFrame to a dictionary
        d = self.df_filtered.to_dict(orient='records')

        # Convert nested JSON data to dictionaries
        for row in d:
            for key, value in row.items():
                if isinstance(value, str):
                    try:
                        row[key] = json.loads(value)
                    except ValueError:
                        row[key] = value
        # Create the JSON object

        json_data = {row['Pattern']: row for row in d}
        # json_data.update(self.main_row)

        # Return the JSON object
        self.json_data = json_data


is_critical_dimension_data_test = 0
if is_critical_dimension_data_test:
    
    ef: str = r'C:\Users\magalang\Documents\examples\DEMO\S32K344-M1AZRAG1.xls'
    cofc_cd = COFCExcelCriticalDimensionWorksheetReader(ef)
    cofc_cd.open_workbook()
    cofc_cd.selected_information = ['Detail']
    cofc_cd.selected_structure_patterns = ['THROUGH_PITCH', 'LINEARITY', 'TLL', 'LES', 'TJP']
    #cofc_cd.selected_structure_patterns = []
    results_data = cofc_cd.create_category_tables()
    cofc_cd.generate_json_data()
    cofc_cd.cd_xml = cofc_cd.critical_dimension_json_to_xml('Critical_Dimension')
    print(cofc_cd.cd_xml)
