######################################################################
# Anes B. - 2017
# A very simple Python CSV to Excel (.xlsx - 2007+) sheet converter.
######################################################################

import csv
import os, glob
import sys
import xlsxwriter

if __name__ == '__main__':
    #listOfFiles = os.listdir(directory)           #  list of all files in the directory
    listOfFiles = glob.glob("*.csv")                       
    for index, fileInList in enumerate(listOfFiles):     
        fileName  = fileInList[0:fileInList.find('.')]     
        excelFile = xlsxwriter.Workbook(fileName + '.xlsx')
        worksheet = excelFile.add_worksheet()    
        #with open(fileName + ".csv", 'rb') as f:
        with open(fileInList, 'rb') as f:   
            content = csv.reader(f)
            for index_row, data_in_row in enumerate(content):
                for index_col, data_in_cell in enumerate(data_in_row):
                    worksheet.write(index_row, index_col, data_in_cell)
    
    excelFile.close()
    print " === Conversion is done ==="