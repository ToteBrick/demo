# -*- coding: utf-8 -*-
import os
import re
from os import path
import sys

sys.path.insert(0, path.dirname(path.dirname(path.abspath(__file__))))
import xlrd
from openpyxl.utils import column_index_from_string


class ExcelResolver:

    def __init__(self, inputFile):
        if os.path.exists(inputFile):
            self.inputFile = inputFile
        else:
            self.inputFile = None

    def get_worksheet_by_index(self, sheet_index=0):
        '''根据index获取工作表'''
        self.sheet_index = int(sheet_index) if isinstance(sheet_index, int) else 0
        if self.inputFile == None:
            print('Excel data file is not exists')
            return False
        else:
            try:
                self.workbook = xlrd.open_workbook(self.inputFile)
                self.worksheet = self.workbook.sheet_by_index(self.sheet_index)
            except Exception as e:
                print('Load excel data file error [%s]' % str(e))
                return False
            else:
                return self.worksheet

    def get_worksheet_by_name(self, sheet_name):
        '''根据名称获取工作表'''
        if self.inputFile == None:
            print('Excel data file is not exists')
            return False
        else:
            try:
                self.workbook = xlrd.open_workbook(self.inputFile)
                self.worksheet = self.workbook.sheet_by_name(sheet_name)
            except Exception as e:
                print('Load excel data file error [%s]' % str(e))
                return False
            else:
                return self.worksheet

    def convert_excel_data_to_dict(self, start_row_num, column_map):

        if not isinstance(column_map, dict):
            return False
        self.xlsx_parse_dicts = []
        for row_num in range(start_row_num, self.worksheet.nrows):
            single_row_dict = {}
            for key, column in column_map.items():
                column_num = column_index_from_string(column) - 1
                cell_value = self.worksheet.row_values(row_num)[column_num]
                single_row_dict[key] = cell_value
            self.xlsx_parse_dicts.append(single_row_dict)

        return self.xlsx_parse_dicts

    def get_all_worksheet(self):
        if self.inputFile == None:
            print('Excel data file is not exists')
            return False
        else:
            try:
                self.workbook = xlrd.open_workbook(self.inputFile)
                sheet_indexs = self.workbook.sheets()
                self.sheet_index_lists = []
                for item in range(len(sheet_indexs)):
                    self.sheet_index_lists.append(self.workbook.sheet_by_index(item))
            except Exception as e:
                print('Load excel data file error [%s]' % str(e))
                return False
            else:
                return self.sheet_index_lists

    def convert_excel_all_data_to_dict(self, start_row_num, column_map):
        if not isinstance(column_map, dict):
            return False
        self.xlsx_all_parse_dicts = []
        for worksheet_info in self.sheet_index_lists:
            for row_num in range(start_row_num, worksheet_info.nrows):
                single_row_dict = {}
                for key,column in column_map.items():
                    column_num = column_index_from_string(column) - 1
                    cell_value = worksheet_info.row_values(row_num)[column_num]
                    single_row_dict[key] = cell_value
                self.xlsx_all_parse_dicts.append(single_row_dict)

        return self.xlsx_all_parse_dicts

