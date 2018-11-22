from pathlib import Path
import openpyxl as pyx
from copy import deepcopy


def findLastRowWithMeaningfulValue(inputSheet):
    # Run through the rows
    for inputRow in range(inputSheet.max_row, 0, -1):
        # Run through the columns
        for inputColumn in range(inputSheet.max_column, 0, -1):
            content = inputSheet.cell(row=inputRow, column=inputColumn).value

            # Return inputRow at the first instance when content is found
            if (str(content) != 'None' and str(content) != ''):
                return inputRow

    return 0  # This happens when the sheet is empty


def makeNewHeadersAndMapColumns(inputSheet, outputSheet, writeParameters, headerDict, columnMappings):

    for col in range(1, inputSheet.max_column + 1):
        columnHeader = inputSheet.cell(row=1, column=col).value

        # add a new header to the dictionary
        if not(columnHeader in headerDict):
            headerDict[columnHeader] = writeParameters['nextWriteColumn']
            outputSheet.cell(
                row=1, column=writeParameters['nextWriteColumn'], value=columnHeader)
            columnMappings[col] = writeParameters['nextWriteColumn']
            writeParameters['nextWriteColumn'] += 1
        else:
            columnMappings[col] = headerDict[columnHeader]

def mergeOneSheet(inputSheet, outputSheet, writeParameters, headerDict):
    columnMappings = {}

    # TODO: Handle cases when the sheet is empty
    lastRow = findLastRowWithMeaningfulValue(inputSheet)

    makeNewHeadersAndMapColumns(inputSheet, outputSheet, writeParameters, headerDict, columnMappings)

    for row in range(2, lastRow + 1):
        isRowEmpty = True
        for col in range(1, inputSheet.max_column + 1):
            content = inputSheet.cell(row=row, column=col).value
            if (str(content) != 'None' and str(content) != ''):
                isRowEmpty = False
            
            outputColumn = columnMappings[col] 
            outputSheet.cell(row=writeParameters['nextWriteRow'], column=outputColumn, value=content)
        
        if (isRowEmpty == False):
            writeParameters['nextWriteRow'] += 1

def mergeSheets(currentDir, stateObject, outputSheet):
    # Standard Excel worksheet format
    # start writing from row 3, row 1 being for the header and row 2 blank
    writeParameters = {
        'nextWriteRow': 3,
        'nextWriteColumn': 1
    }

    headerDict = {}

    # For every file, run through their sheets
    for inputFile, inputSheets in stateObject['fileAndSheetDict'].items():

        fileDir = currentDir / 'Spreadsheets' / inputFile
        sourceWorkbook = pyx.load_workbook(fileDir, data_only=True)

        # For every sheet to merge
        for sheet in inputSheets.split(','):
            if sheet in stateObject['dontMerge']:
                continue

            # Putting the contents of the current sheet into the output sheet
            mergeOneSheet(sourceWorkbook[sheet],
                          outputSheet, writeParameters, headerDict)
