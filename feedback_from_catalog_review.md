# Notes from Catalog Reviews

- The header isn't always on the first row
  - Maybe take the header row as an input and the start of the content row as an input too
- There are some dummy rows at the end in the source which have no meaningful content
  - You have to take the end row as input too it seems
- Some rows are just category headings (aka useless data).
  - Maybe you need to have an option in the config where you say you skip certain rows.
- Some catalogs have the same data spread across multiple sheets
  - Maybe you could also have an option to enter the start row and column of the output excel sheet
  - Then, multiple configs could write to the same file
  - Then you could also take the name of the output file. If it doesn't exist, then generate it.
 
