import unittest
import xlsxwriter
import os
import xlrd
from openpyxl import load_workbook, Workbook

class ExcelRead:
    def xlrd_read_excel(self, path_file):
        data_return = ()
        book = xlrd.open_workbook(path_file)
        ws = book.sheet_by_index(0)
        raw_headers = [cell.value for cell in ws.row(0)]
        col_count = len(raw_headers)
        for rx in range(ws.nrows):
            data_return += ([ws.cell_value(rowx=rx, colx=ci) for ci in range(col_count)], )
        return data_return

    def openpyxl_read_execl(self, path_file):
        wb = load_workbook(path_file)
        sheet = wb.active
        data_return = ()
        max_row = sheet.max_row
        max_column = sheet.max_column
        # iterate over all cells
        for i in range(1, max_row + 1):
            row = []
            for j in range(1, max_column + 1):
                cell_obj = sheet.cell(row=i, column=j)
                row.append(cell_obj.value)
            data_return += (row,)
        return data_return

class ExcelWrite:
    def xlsxwriter_write_to_excel(self, file_name, worksheet_name, data_to_write=()):
        # Create an new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet(worksheet_name)  # Defaults to Sheet1

        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})
        # Add a number format for cells with money.
        money = workbook.add_format({'num_format': '$#,##0'})
        # Some data we want to write to the worksheet.
        expenses = data_to_write
        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0
        # Iterate over the data and write it out row by row.
        for item, cost in (expenses):
            worksheet.write(row, col, item)
            worksheet.write(row, col + 1, cost)
            row += 1
        workbook.close()

    def openpyxl_write_excel(self, file_name, data_to_write=()):
        wb = Workbook()
        sheet = wb.active
        # append all rows
        for row in data_to_write:
            sheet.append(tuple(row))
        # save file
        wb.save(file_name)

class TestExcelRun(unittest.TestCase):

    def test_xlrd_read_excel(self):
        # create fixture
        expenses = (['Rent', 1000], ['Gym', 50])
        excel_file_name = 'test.xlsx'
        excelread = ExcelRead()

        # check excel file is exist and it's data
        assert os.path.isfile(excel_file_name)
        assert excelread.xlrd_read_excel(excel_file_name) == expenses

    def test_openpyxl_read_execl(self):
        # create fixture
        expenses = (['Rent', 1000], ['Gym', 50])
        excel_file_name = 'test.xlsx'
        excelread = ExcelRead()

        # check excel file is exist and it's data
        assert os.path.isfile(excel_file_name)
        assert excelread.openpyxl_read_execl(excel_file_name) == expenses

    def test_xlsxwriter_write_to_excel(self):
        # create fixture
        expenses = (['Rent', 1000], ['Gym', 50])
        excel_file_name, worksheet_name = 'test2.xlsx', 'Sheet test'
        excelread, excelwrite= ExcelRead(), ExcelWrite()
        excelwrite.xlsxwriter_write_to_excel(excel_file_name, worksheet_name, expenses)

        # check excel file is exist and it's data
        assert os.path.isfile(excel_file_name)
        assert excelread.xlrd_read_excel(excel_file_name) == expenses

    def test_openpyxl_write_to_excel(self):
        # create fixture
        expenses = (['Rent', 1000], ['Gym', 50])
        excel_file_name, worksheet_name = 'test2.xlsx', 'Sheet test'
        excelread, excelwrite= ExcelRead(), ExcelWrite()
        excelwrite.openpyxl_write_excel(excel_file_name,  expenses)

        # check excel file is exist and it's data
        assert os.path.isfile(excel_file_name)
        assert excelread.openpyxl_read_execl(excel_file_name) == expenses
