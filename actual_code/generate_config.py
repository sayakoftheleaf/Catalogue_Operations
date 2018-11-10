import toml
import openpyxl

sourceFile = input("Please enter the name of the source Excel document: ")
sourceSheet = input(
    "Please enter the name of the sheet you want to be mapped: ")
numberOfBoxes = input(
    "Please enter the number of boxes that need to be shipped: ")
configFileName = input(
    "Please enter the name of the config file to be generated: ")

workBook = openpyxl.load_workbook(sourceFile, read_only=True)
workSheet = workBook[sourceSheet]

workRows = workSheet.max_rows
workColumns = workSheet.max_columns

# create the toml file
fileName = configFileName + ".toml"
file = open(fileName, w+)
file.close()

# create the sheet data and dump it into the toml file
sheetDict = {
  'name': sourceSheet,
  'rows': workRows,
  'cols': openpyxl.utils.get_column_letter(workColumns)
}

toml.dumps(sheetDict, fileName)


dataDict = {}
sourceHeaderDict = {}
outputHeaderDict = {}

dataDict['filename'] = sourceFile

# Extract total number of rows from sheet
# totalRowsInSheet =

# Extract total number of columns from sheet
# totalColsSheet =

# Convert that into a Column Letter
# totalCoslSheet =

# Map through the first row and read the contents of the cell


def generateSingleHeader(columnRef, header):
    return {
        columnRef: '',
    }

# def generateContentBeforeBoxes():

# def generateContentForBoxes():

# def generateContentAfterBoxes():
