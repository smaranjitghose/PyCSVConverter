import xlrd  # This package is used to extract data from Excel spreadsheets
import csv  # This module implements classes to read and write tabular data in CSV format

def csv_from_excel(xlsx_file_path,csv_file_path,sheet_name='Sheet1'):
    #Opening the xlxs file
    wb = xlrd.open_workbook(xlsx_file_path)
    #Accessing the desired sheet of the xlxs file
    sheet = wb.sheet_by_name(sheet_name)
    #Creating a csv file to write the data from xlxs file
    csv_file = open(csv_file_path, 'w')
    wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
    #Writing the values row wise into the csv file
    for rownum in range(sheet.nrows):
        wr.writerow(sheet.row_values(rownum))
    #Closing the xlxs file
    csv_file.close()


# Taking in the path of the xlsx file, the sheet to be accesed and the path of the output csv file
xlsx_file_path = input("Enter the path of the xlsx file: ")
sheet_name = input("Enter the name  of the sheet: ")
csv_file_path = input("Enter the path of the csv file: ")
csv_from_excel(xlsx_file_path, csv_file_path, sheet_name)
