# TODO: Implement cases where one file maps to multiple sources

def writeMaps(fileObject, sheet, lastCol):
    commentString = "\n# Put all the mappings here\n# If one column maps to several columns, separate them with a comma\n# DO NOT PUT WHITESPACE AFTER COMMA!\n# For example: mapsto.A = \"B,X\"\n\n"
    colString = commentString
    # iterate through all of the columns
    for currentCol in range(1, lastCol + 1):
        # get the letter equivalent of the current column
        columnLetter = openpyxl.utils.get_column_letter(currentCol)
        # get the header inside the cell
        header = sheet.cell(row=1, column=currentCol).value
        colString += "mapsto.{0} = \"\" # {1}\n".format(columnLetter, header)

    fileObject.write(colString)


def evaluateOrder(boxInformationOrder):
    # Removing commas and spaces and converting everything to uppercase
    boxInformationOrder = boxInformationOrder.replace(" ", "")
    boxInformationOrder = boxInformationOrder.replace(",", "")
    boxInformationOrder = boxInformationOrder.upper()

    orderDict = {}
    for value in ['L', 'B', 'H', 'W']:
        orderDict[value] = boxInformationOrder.index(value)

    return orderDict

# FIXME:
def mapBoxes (fileObject, sheet, sourceBoxesStartFrom, boxInformationOrder):
    orderDict = evaluateOrder(boxInformationOrder)
