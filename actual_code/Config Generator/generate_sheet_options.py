import toml
import openpyxl

def writeSheetOptions(fileObject, name, rows, cols):
    commentString = "\n# Details of the source sheet. Edit in case info is wrong \n\n"
    name = "sheet.name = \"{0}\"\n".format(name)
    rows = "sheet.rows = +{0}\n".format(str(rows))
    cols = "sheet.cols = \"{0}\"\n".format(
        openpyxl.utils.get_column_letter(cols))

    fileObject.write("{0}{1}{2}{3}".format(commentString, name, rows, cols))