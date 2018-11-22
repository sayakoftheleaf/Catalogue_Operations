import openpyxl

def mapHeaders (configDict, outputSheet):
  # placing all the headers
  for key, value in configDict.get('header').items():
    columnIndex = openpyxl.utils.column_index_from_string(str(key))
    outputSheet.cell(row=1, column=columnIndex, value=str(value))

  return outputSheet