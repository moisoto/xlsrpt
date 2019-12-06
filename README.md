# xlsrpt
Excel Reports Generator Package

### Overview
Easy & Flexible Excel Report Generator Package for go.

### Features
- Generate simple Excel Report by providing *sql.DB object and a Query String
- Generate more flexible, more complex Excel Report by providing a map of structures where each struct element is a record.
- Struct Based generation can specify column names, columns to be summarized, allows specific order using unique columns.

### Dependencies
This package currently depends on [tealeg's xlsx](https://github.com/tealeg/xlsx) v1.0.5 package. 
xlsrpt uses go modules, no need to worry if you tealeg/xlsx in your programs along with this package.

## Project State
xlsrpt is currently in beta, all main functionality is completed and working.
Pending work before first release:
- Unit testing code
- Code cleanup / verification
- Commented code cleanup
- Better documentation
