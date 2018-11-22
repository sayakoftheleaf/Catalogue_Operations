
import openpyxl as pyx
from pathlib import Path

from MergeSheets import mergeSheets
from HandleInput import acceptInputAndFormFileDict, correctExtenstion

stateObject = {
    'dontMerge': [],
    'fileAndSheetDict': {}
}

def main():
    global stateObject

    currentDir = Path('./..')

    # accept the Input Files and Sheets
    acceptInputAndFormFileDict(currentDir, stateObject)

    outputFile = input('Enter the name of the output file to be generated: ')
    outputSheet = input('Enter the name of the output sheet to be generated: ')

    # auto-correcting the output extension and resolving the Path
    outputFile = currentDir / 'Generated' / correctExtenstion(outputFile)

    outputWorkbook = pyx.Workbook()
    writeSheet = outputWorkbook.create_sheet(title=outputSheet)

    mergeSheets(currentDir, stateObject, writeSheet)

    # trimWhiteSpaceFromOutput(writeSheet)
    outputWorkbook.save(outputFile)

    print('Sheets have been merged and output file has been generated!')


if __name__ == "__main__":
    main()
