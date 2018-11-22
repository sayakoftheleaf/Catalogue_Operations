# Catalogue_Operations

**Sheet Mapper** - Maps an .xlsx document into another .xlsx document given a .toml config file
**Config Generator** - Generates the outline of a config that can then be modified to be used in the Sheet Mapper
**Sheet Merger** - Merges multiple sheets into one sheet. Can operate across mulitple .xlsx files and all .xlsx files on a subdirectory.

**NOTE:** .xls format is not supported currently

## Setup Instructions

- Create a directory 'Configs' to save configs generated
- Create a directory 'Spreadsheets' to save all of the raw spreadsheet files to be worked on
  - In case you want to operate on a subdirectory, the subdirectory must also go here
- Create a directory 'Generated' where the generated Spreadsheets are going to be stored
