#!/usr/bin/env python
# coding: utf-8

import xml.etree.ElementTree as ET
from math import isnan
from typing import Dict, Any

import pandas as pd
import numpy as np
import xml.dom.minidom as minidom
import json
import openpyxl

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)


class COFCExcelStatisticsWorksheetReader:
    main_row: Dict[Any, Any]

    def __init__(self, file_path):
        self.file_path: str = file_path
        self.workbook: str = 'Statistics'
        self.current_sheet = None
        self.main_row = {}
        self.df = None
        self.table = None
        self.transformed_data = None
        self.patterns = None
        self.json_data = None

    def open_workbook(self, workbook="Statistics"):
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

    def transform_stat_tables(self, tab, table):

        tab = tab.strip()
        tab = tab.replace(' $', '')
        tab = tab.replace('.', '')
        tab = tab.replace('/', '_')
        tab = tab.replace(' ', '_')
        tab = tab.replace('\n', '_')
        table = table.iloc[1:].reset_index(drop=True)
        patterns = list(table.iloc[0, 0:])
        patterns = [v.strip() for v in patterns]
        patterns = [v.replace(' $', '') for v in patterns]
        patterns = [v.replace('.', '') for v in patterns]
        patterns = [v.replace('/', '_') for v in patterns]
        patterns = [v.replace(' ', '_') for v in patterns]
        patterns = [v.replace('\n', '_') for v in patterns]
        table = table.iloc[1:].reset_index(drop=True)
        table.columns = patterns

        if tab == 'Registration':
            table['Ortho'] = table['Ortho'].fillna(method='ffill')
        if tab == 'Critical Dimension':
            table['X_Y_Diff'] = table['X_Y_Diff'].fillna(method='ffill')
        table.reset_index(drop=True)
        table['Category'] = tab
        transformed_table = table.apply(lambda row: json.dumps(row.to_dict()), axis=1).tolist()
        return transformed_table

    def process_table_groups(self):
        df = self.df
        df = df.dropna(axis=1, how='all')

        # Get the patterns as column names
        max_row = 0
        row = 0
        for keys in ['Device', 'Layer', 'Lot No.']:
            try:
                row = df[df[1] == keys].index[0]
                self.main_row.update({keys: str(df.loc[row][2])})
                if max_row < row:
                    max_row = row
            except Exception as e:
                continue
        df = df.drop(df.index[:max_row+1])
        df = df.reset_index(drop=True)
        nan_rows = df.isna().all(axis=1)
        idx_lists = []
        true_idx = []
        for index, value in nan_rows.items():
            if not value:
                true_idx.append(index)
            elif true_idx:
                idx_lists.append(true_idx)
                true_idx = []
        if len(true_idx) > 0:
            idx_lists.append(true_idx)

        tables = {}
        for idxl in idx_lists:
            table = df.iloc[idxl[0]:idxl[-1]+1].dropna(axis=1, how='all')
            # Replace nan values in column 1 with category patterns.
            table[1] = table[1].fillna(method='ffill')
            if not table.empty:
                tt = table.iloc[0, 0]

                if isinstance(table.iloc[0,0], float):
                    tt = 'Critical Dimension'
                if tt in tables:
                    tables[tt].append(table)
                else:
                    tables[tt] = [table]

        transformed_data = []
        for tab in tables:
            for table in tables[tab]:
                try:
                    transformed_data.append(self.transform_stat_tables(tab, table))
                except Exception as e:
                    print(f"Transform Stat Table Error: {str(e)}")
        self.transformed_data = transformed_data
        return transformed_data

    def statistics_json_to_xml(self, root_name='Statistics'):
        transformed_data = self.transformed_data
        root = ET.Element(root_name)
        for tabs in transformed_data:
            for x in tabs:
                dx = json.loads(x)
                category = dx['Category']
                site_element = ET.SubElement(root,category)
                for key, value in dx.items():
                    if key == 'Category':
                        continue
                    if key == 'Phase_Trans':
                        continue
                    key = key.replace(' ', '_')
                    key = key.replace('-', '_')
                    if key.lower() == '3_sigma' or key.lower() == '3sigma':
                        key = 'Stat_3_sigma'
                    child_elem = ET.SubElement(site_element, key)
                    try:
                        if value == 'N/A':
                            value = 'N_A'
                        ET.fromstring(f"<root>{value}</root>")
                        child_elem.text = str(value)
                    except ET.ParseError:
                        print("Errored:",value)
                        exit()

        xml_string = ET.tostring(root, encoding='unicode')
        xml_dom = minidom.parseString(xml_string)
        pretty_xml = xml_dom.toprettyxml(indent='  ')
        return pretty_xml

    def generate_json_data(self):
        # No need.
        pass

    def get_worksheets(self):
        workbook = openpyxl.load_workbook(self.file_path)
        sheet_names = workbook.sheetnames
        return sheet_names



is_statistics_data_test = 0
if is_statistics_data_test:
    ef: str = r'C:\Users\magalang\Documents\\COA_DEVICE_SAMP.protected.xls'
    cofc_stat = COFCExcelStatisticsWorksheetReader(ef)
    cofc_stat.open_workbook()
    results_data = cofc_stat.process_table_groups()
    cofc_stat.generate_json_data()
    cofc_stat.cd_xml = cofc_stat.statistics_json_to_xml('Statistics')
    print(cofc_stat.cd_xml)
