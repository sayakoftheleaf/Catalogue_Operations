import openpyxl
from pathlib import Path

# new dictionary that has headers as keys and the column they output
# to in the output file as their values
fileAndSheetDict = {}
headerDict = {}


def correctExtenstion(someString):
    if (len(someString.split('.')) != 2 or someString.split[1] != 'xlsx'):
        return someString.split('.')[0] + '.xlsx'
    else:
        return someString

# Writes all the contents of a column into the output Sheet


def readAndOutputColumn(inputColumn, inputSheet, outputSheet, maxRow, writeRow, nextWriteColumn):
    # we need to modify global Dict
    global headerDict

    # assuming that headers are in the first row of the input sheet
    columnHeader = inputSheet.cell(row=1, column=inputColumn).value

    # add a new header to the dictionary
    if not(columnHeader in headerDict):
        headerDict[columnHeader] = nextWriteColumn
        outputSheet.cell(row=1, column=nextWriteColumn, value=columnHeader)
        nextWriteColumn += 1

    outputColumn = headerDict[columnHeader]

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
    currentSheet = inputWorkbook[sheetName]

    # for every column in the current sheet
    for currentColumn in range(1, currentSheet.max_column + 1):
        nextWriteColumn = readAndOutputColumn(
            currentColumn, currentSheet, outputSheet, currentSheet.max_row, nextWriteRow, nextWriteColumn)

    # the next sheet needs to print content below the current sheet
    nextWriteRow += currentSheet.max_row + 1

    return [nextWriteRow, nextWriteColumn]


def acceptInputAndFormFileDict():
    global fileAndSheetDict

    numberOfFilesToMerge = input(
        'Enter the number of excel files you need to merge: ')

    for i in range(0, int(numberOfFilesToMerge)):
        fileNumber = str(i + 1)
        fileName = input(
            'Enter the name of file {0} :'.format(fileNumber))
        inputSheets = input(
            'Enter the names of the sheets in file {0}'.format(fileNumber)
            + 'to be merged (multiple files can be separated by commas): ')

        # removing whitespace
        inputSheets = inputSheets.replace(', ', ',')

        # correcting the input Name
        fileName = correctExtenstion(fileName)
        fileAndSheetDict[fileName] = inputSheets

def main():

    currentDir = Path('./..')

    # accept the Input Files and Sheets
    acceptInputAndFormFileDict()

    outputFile = input('Enter the name of the output file to be generated: ')
    outputSheet = input('Enter the name of the output sheet to be generated: ')

    # auto-correcting the output extension and resolving the Path
    outputFile = currentDir / 'Generated' / correctExtenstion(outputFile)


    outputWorkbook = openpyxl.Workbook()
    writeSheet = outputWorkbook.create_sheet(title=outputSheet)

    # next possible column a header can be written in
    # in the output sheet
    nextWriteColumn = 1

    # next row in the output where the content can be written
    nextWriteRow = 3

    # For every file, run through their sheets
    for inputFile, inputSheets in fileAndSheetDict:
        
        fileDir = currentDir / 'Spreadsheets' / inputFile
        sourceWorkbook = openpyxl.load_workbook(fileDir, data_only=True)

        # For every sheet to merge
        for sheet in inputSheets.split(','):
            # TODO: There has to be a better way to do this
            temp = mergeSheet(sourceWorkbook, sheet, writeSheet,
                          nextWriteRow, nextWriteColumn)
            nextWriteRow = temp[0]
            nextWriteColumn = temp[1]

    outputWorkbook.save(outputFile)
    print('Sheets have been merged and output file has been generated!')

if __name__ == "__main__":
    main()
