
# Project Structure - Data Concatenation and Transformation module

## IO

- Input - <source>.xlsx <config>.toml
- Output - <output>.xlsx

## Assumptions 

  - All data content are valid strings (i.e strings with valid unicode characters/ empty strings)
  - The .xlsx document only has one sheet with all the data

- FUTURE TODO: Add other data type compatibility and type checking

## Program Outline

- Step 1 - Separate Program - Generate a skeleton config file that the user can edit. (Also can be just copied from a manually created skeleton)
- Step 2 - Accept the source file and the config file as input
- Step 3 - Loop through the contents of the file
  - Extract row and column number from the config file
  - External loop: rows 
  - Internal loop: columns
  - For each column, 
    - extract the contents of the cell into a String
    - check the config file for behavior, and map to a dictionary
      - Possible behaviors : 
        - directly maps to a new cell in the output
        - doesn't map to anything in the output
        - maps to an existing cell, in which case, separate the previous content with the new content with the separator defined in the config file
- Step 4 - convert the dictionary into a new .xlsx document



