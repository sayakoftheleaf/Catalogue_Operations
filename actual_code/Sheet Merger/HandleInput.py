from os import listdir
import openpyxl as pyx

# A bunch of verification checks prompted to the user to ensure correct
# input file format


def fileChecks():
    sheetVerification = input(
        'Are all the sheets in the files to be merged? (y for YES or n for NO): ')
    if (sheetVerification == 'n' or sheetVerification == 'N'):
        print('Please delete the unnecessary sheets and try again.')
        exit()  # Currently cannot handle unnecessary sheets

    xlsxVerification = input(
        'Are all the files in .xlsx format? (.xls is not supported) (y or n): ')
    if (xlsxVerification == 'n' or xlsxVerification == 'N'):
        print('Please convert the files to .xlsx and try again.')
        exit()

    headerVerification = input(
        'Do all the sheets have headers ONLY on Row 1 and no other row? (y or n): ')
    if (headerVerification == 'n' or headerVerification == 'N'):
        print('Please modify the headers and put them on the first row before running this program.')
        exit()

# Add an extension to the filename entered as input if it already doesn't have one
# Or correct a wrong extension hoping it was a mistype


def correctExtenstion(someString):
    if (len(someString.split('.')) != 2 or someString.split[1] != 'xlsx'):
        return someString.split('.')[0] + '.xlsx'
    else:
        return someString

# Create a dictionary of the files and the sheets in them that need to be merged


def addFileToDict(fileName, sheetNames, fileAndSheetDict):

    fileAndSheetDict[fileName] = sheetNames


def acceptFilesFromDirectory(currentDir, dontMerge, fileAndSheetDict):

    dirName = input(
        'Please enter the name of the directory that has the files to be merged: ')

    dontMerge.append(input('enter sheet names that will be skipped: '))

    fileChecks()

    for currentFile in listdir(currentDir / 'Spreadsheets' / dirName):
        extension = currentFile.split('.')
        extension = extension[len(extension) - 1]

        if (extension == 'xlsx'):
            fullPath = currentDir / 'Spreadsheets' / dirName / currentFile
            currentWorkbook = pyx.load_workbook(fullPath)

            sheets = ','.join(currentWorkbook.sheetnames)

            addFileToDict(dirName + '/' + currentFile,
                          sheets, fileAndSheetDict)


def acceptMultipleFiles(dontMerge, fileAndSheetDict):
    fileChecks()

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

        addFileToDict(fileName, inputSheets, fileAndSheetDict)


def acceptInputAndFormFileDict(curentDir, stateObject):

    navigationType = input(
        'Are you going to merge individual files'
        'or all files in a directory (f for file, d for dir): ')

    if (navigationType == 'd'):
        acceptFilesFromDirectory(
            curentDir, stateObject['dontMerge'], stateObject['fileAndSheetDict'])
    elif (navigationType == 'f'):
        acceptMultipleFiles(
            stateObject['dontMerge'], stateObject['fileAndSheetDict'])
    else:
        print('Invalid Input. Please try again.')
        acceptInputAndFormFileDict(curentDir, stateObject)
