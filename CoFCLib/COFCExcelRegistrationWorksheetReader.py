#!/usr/bin/env python
# coding: utf-8

import xml.etree.ElementTree as ET
from math import isnan
from typing import List, Any

import pandas as pd
from lxml import etree

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

class COFCExcelRegistrationWorksheetReader:
    def __init__(self, file_path):
        self.file_path = file_path
        self.workbook = 'Registration'
        self.current_sheet = None
        self.df = None
        self.registration_data = []
        self.main_patterns = []
        self.registration_data_table = []
        self.patterns = []
        self.registration_dict_data = []
        self.main_row = {}
        self.reg_xml = ''

    def open_workbook(self, workbook="Registration"):
        """
        :type workbook:str
        """
        self.workbook = workbook
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=self.workbook, header=None, usecols=None,
                                    keep_default_na=False, na_values=['','nan','NaN'])
        except ValueError as e:
            print("e")
            exit(1)
        except Exception as e:
            print("e")
            exit(1)

    def generate_registration_data(self):
        # Read Excel file and create category tables
        self.create_category_tables()

        self.registration_dict_data = {}
        # combine the "Specifications" rows
        try:
            new_spec_row = self.transform_specification_rows()
            self.registration_dict_data['Specifications'] = new_spec_row
        except Exception as e:
            print(f"Specifications Row Error: {str(e)}")

        # combine the "Result" rows
        try:
            new_result_row = self.transform_results_rows()
            self.registration_dict_data['Result'] = new_result_row
        except Exception as e:
            print(f"Result Row Error: {str(e)}")

        # combine the "Detail" rows

        try:
            new_detail_row = self.transform_details_row()
            self.registration_dict_data['Detail'] = new_detail_row
        except Exception as e:
            print(f"Detail Row Error: {str(e)}")

        return

    def create_category_tables(self):
        df = self.df
        # Get the patterns as column names
        self.main_row = {}
        for keys in ['Device', 'Layer', 'Tool']:
            row = df[df[1] == keys].index[0]
            self.main_row.update({keys: str(df.loc[row][2])})

        start_row = df[df[1].str.startswith('Information', na=False)].index[0]
        start_col = df.columns.get_loc(1)
        # Replace NaN values in Categories with the category name
        patterns = list(df.iloc[start_row, start_col + 1:])
        self.main_patterns = [x for x in patterns if str(x) != 'nan']
        patt_xyo = list(df.iloc[start_row + 1, start_col + 1:])
        new_pattern = []
        non_nan_val = 'SubCategory'
        for i in range(len(patterns)):
            if isinstance(patterns[i], str):
                non_nan_val = patterns[i]
            if isinstance(patterns[i], float) and isnan(patterns[i]):
                patterns[i] = non_nan_val
            new_pattern.append(str(patterns[i]) + ',' + str(patt_xyo[i]))

        patterns = new_pattern
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
        column_names = ['Categories'] + patterns

        # Slice the table from the start row and column
        table = df.iloc[start_row + 1:, start_col:]

        # Rename columns
        table.columns = column_names

        # Replace NaN values in Categories with the category name
        table['Categories'] = table['Categories'].fillna(method='ffill')

        # Drop the first row of the table
        table = table.iloc[1:].reset_index(drop=True)

        # Reset the index to ensure all values are unique
        table = table.reset_index(drop=True)
        # print(table)

        self.registration_data_table = table
        self.patterns = patterns

    def transform_details_row(self):
        full_table: List[Any] = self.registration_data_table
        # patterns = self.patterns
        reg_names: List[Any] = self.main_patterns

        # Filter the table to get only rows with "Categories" == "Detail"
        table = full_table[full_table['Categories'] == 'Detail']
        new_results = {}
        try:
            for pat in reg_names:
                newcolumns = [col for col in table.columns if col.startswith(pat + ',')]
                new_colname = [col[len(pat + ','):] for col in newcolumns]
                new_pat_table = table[newcolumns].rename(columns=dict(zip(newcolumns, new_colname)))
                new_pat_table = new_pat_table.dropna(how='all')
                new_pat_table[pat] = new_pat_table.apply(lambda row: {col: row[col] for col in new_pat_table.columns},
                                                         axis=1)
                new_results[pat] = new_pat_table[pat].reset_index(drop=True).tolist()
                new_pat_table = None
        except Exception as e:
            print(f"Error occurred while combining result rows: {e}")
        return new_results

    def transform_results_rows(self):
        full_table: List[Any] = self.registration_data_table
        # patterns = self.patterns
        reg_names: List[Any] = self.main_patterns

        # Filter the table to get only rows with "Categories" == "Result"
        table = full_table[full_table['Categories'] == 'Result']
        new_results = {}
        try:
            for pat in reg_names:
                newcolumns = [col for col in table.columns if col.startswith(pat + ',')]
                new_colname = [col[len(pat + ','):] for col in newcolumns]
                new_pat_table = table[newcolumns].rename(columns=dict(zip(newcolumns, new_colname)))
                new_pat_table.insert(0, 'Type', table['SubCategory,X/Y'])
                new_pat_table[pat] = new_pat_table.apply(lambda row: {col: row[col] for col in new_pat_table.columns},
                                                         axis=1)
                new_results[pat] = new_pat_table[pat].reset_index(drop=True).tolist()
                new_pat_table = None
        except Exception as e:
            print(f"Error occurred while combining result rows: {e}")
        return new_results

    def transform_specification_rows(self):
        full_table: List[Any] = self.registration_data_table
        # patterns = self.patterns
        reg_names: List[Any] = self.main_patterns

        # Filter the table to get only rows with "Categories" == "Specification"
        table = full_table[full_table['Categories'] == 'Specification']
        new_results = {}
        try:
            for pat in reg_names:
                newcolumns = [col for col in table.columns if col.startswith(pat + ',')]
                new_colname = [col[len(pat + ','):] for col in newcolumns]
                new_pat_table = table[newcolumns].rename(columns=dict(zip(newcolumns, new_colname)))
                new_pat_table.insert(0, 'Type', table['SubCategory,X/Y'])
                new_pat_table[pat] = new_pat_table.apply(lambda row: {col: row[col] for col in new_pat_table.columns},
                                                         axis=1)
                new_results[pat] = new_pat_table[pat].reset_index(drop=True).tolist()
                new_pat_table = None
        except Exception as e:
            print(f"Error occurred while combining specification rows: {e}")
        return new_results

    def regdata_to_xml(self, root_name='Root'):
        regdata = self.registration_dict_data
        # Create the root element of the XML
        root = ET.Element(root_name)
        for row in regdata:
            child_element = ET.SubElement(root, row)
            df = regdata[row]
            # Iterate over the dataframe rows
            for name, values in df.items():
                site_element = ET.SubElement(child_element, 'REG_Site')
                name_element = ET.SubElement(site_element, 'Name')
                name_element.text = name
                for info in values:
                    info_element = ET.SubElement(site_element, 'Info')
                    for key, value in info.items():
                        info_element.set(key, str(value))
        # Create the XML tree
        xml_string = ET.tostring(root, encoding='unicode')
        x = etree.fromstring(xml_string)
        return etree.tostring(x, pretty_print=True, encoding=str)

    def main_reader_sample(self):
        self.open_workbook()
        self.generate_registration_data()
        self.reg_xml = self.regdata_to_xml('Registration')


"""
    TEST!!!
"""
is_registration_data_test = 0
if is_registration_data_test:
    ef: str = r'C:\Users\magalang\Documents\COA_DEMO2.protected.xls'
    cofc_reg = COFCExcelRegistrationWorksheetReader(ef)
    cofc_reg.open_workbook()
    cofc_reg.generate_registration_data()
    # cofc_reg.main_test_data_to_xml()
    cofc_reg.reg_xml = cofc_reg.regdata_to_xml('Registration')
    print(cofc_reg.reg_xml)
