import openpyxl
from pathlib import Path

def main():
    currentDir = Path('./..')

    fileDir = input('Enter name of the input xlsx file: ')
    outputFile = input('Enter the name of the output file to be generated: ')
    inputSheets = input(
        'Enter the names of the sheets to be merged (multiple files can be separated by commas): ')
    outputSheet = input('Enter the name of the output sheet to be generated: ')
    
    # auto-correcting the input extension
    if (len(fileDir.split(.)) != 2 or fileDir.split[1] != 'xlsx'):
        fileDir = fileDir.split('.')[0] + '.xlsx'
      
    # auto-correcting the output extension
    if (len(outputFile.split(.)) != 2 or outputFile.split[1] != 'xlsx'):
        outputFile = outputFile.split('.')[0] + '.xlsx'

    # removing whitespace
    inputSheets = inputSheets.replace(' ', '')

    # resolving Paths
    fileDir = currentDir / 'Spreadsheets' / fileDir
    outputFile = currentDir / 'Generated' / outputFile

    sourceWorkbook = openpyxl.load_workbook(fileDir, data_only = True)
    outputWorkbook = openpyxl.Workbook()
    writeSheet = outputWorkbook.create_sheet(title = outputSheet)

    sheetDict = {}

    for sheet in inputSheets.split(','):
        currentWorkbook = sourceWorkbook[sheet]
        sheetColumns = currentWorkbook.max_column
        sheetRows = currentWorkbook.max_row

        for currentColumn in range (1, sheetColumns + 1):
            columnHeader = currentWorkbook.cell(row=1, column=currentColumn)
            if header in sheetDict:
                

            



if __name__ == "__main__":
    main()