import openpyxl
from pathlib import Path

# new dictionary that has headers as keys and the column they output
# to in the output file as their values
headerDict = {}


def correctExtenstion(someString):
    if (len(someString.split(.)) != 2 or someString.split[1] != 'xlsx'):
        return someString.split('.')[0] + '.xlsx'
    else:
        return someString

# Writes all the contents of a column into the output Sheet


def readAndOutputColumn(inputColumn, inputSheet, outputSheet, maxRow, writeRow, nextWriteColumn):
    # we need to modify global Dict
    global headerDict

    # assuming that headers are in the first row of the input sheet
    columnHeader = inputSheet.cell(row=1, column=inputColumn)

    # add a new header to the dictionary
    if not(columnHeader in headerDict):
        headerDict[columnHeader] = nextWriteColumn
        nextWriteColumn += 1

    outputColumn = sheetDict[columnHeader]

    # write all the contents of that column in the
    # corresponding output sheet
    # Assuming that the content in the source sheet begin from row 3
    for currentRow in range(3, maxRow + 1):
        content = inputSheet.cell(row=currentRow, column=inputColumn).value
        outputSheet.cell(row=writeRow,
                         column=outputColumn, value=str(content))
        writeRow += 1
    
    return nextWriteColumn


def mergeSheet(inputWorkbook, sheetName, outputSheet, nextWriteRow, nextWriteColumn):
    # load sheet
    currentSheet = inputWorkbook[sheet]

    # for every column in the current sheet
    for currentColumn in range(1, currentSheet.max_column + 1):
        nextWriteColumn = readAndOutputColumn(
            currentColumn, currentSheet, outputSheet, currentSheet.max_row, nextWriteRow, nextWriteColumn)

    # the next sheet needs to print content below the current sheet
    nextWriteRow += currentSheet.max_row + 1

    return [nextWriteRow, nextWriteColumn]


def main():
    currentDir = Path('./..')

    fileDir = input('Enter name of the input xlsx file: ')
    outputFile = input('Enter the name of the output file to be generated: ')
    inputSheets = input(
        'Enter the names of the sheets to be merged (multiple files can be separated by commas): ')
    outputSheet = input('Enter the name of the output sheet to be generated: ')

    # auto-correcting the input and output extensions
    fileDir = correctExtenstion(fileDir)
    outputFile = correctExtenstion(outputFile)

    # removing whitespace
    inputSheets = inputSheets.replace(' ', '')

    # resolving Paths
    fileDir = currentDir / 'Spreadsheets' / fileDir
    outputFile = currentDir / 'Generated' / outputFile

    sourceWorkbook = openpyxl.load_workbook(fileDir, data_only=True)
    outputWorkbook = openpyxl.Workbook()
    writeSheet = outputWorkbook.create_sheet(title=outputSheet)

    # next possible column a header can be written in
    # in the output sheet
    nextWriteColumn = 1

    # next row in the output where the content can be written
    nextWriteRow = 3

    # For every sheet to merge
    for sheet in inputSheets.split(','):
        # TODO: There has to be a better way to do this
        temp = mergeSheet(sourceWorkbook, sheet, writeSheet, nextWriteRow, nextWriteColumn)
        nextWriteRow = temp[0]
        nextWriteColumn = temp[1]

if __name__ == "__main__":
    main()
