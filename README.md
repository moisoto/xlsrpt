[![made-with-Go](https://img.shields.io/badge/Made%20with-Go-1f425f.svg?color=%237fd5ea)](http://golang.org)
[![GoReportCard](https://goreportcard.com/badge/github.com/moisoto/xlsrpt)](https://goreportcard.com/report/github.com/moisoto/xlsrpt)

# xlsrpt
Excel Reports Generator Package

### Overview
Easy & Flexible Excel Report Generator Package for go.

### Usage
#### For detailed documentation and examples please see https://godoc.org/github.com/moisoto/xlsrpt

There are four ways to generate a report. Each of the following functions can be used depending your needs:
- ExcelFromDB()
  - Allows generation of report with minimal effort. Uses reflect to infer column types.

- ExcelReport()
  - Allows single sheet report generation. You must define a struct with fields for each column.
  - A function that loads data into a map of such struct must be implemented.

- ExcelMultiSheetFromDB()
  - Allows multiple sheets report generation with minimal effort. Uses reflect to infer column types.

- ExcelMultiSheet()
  - Allows multiple sheets report generation. 
  - You must define a struct for each sheet, each one with fields for each column to be added. 
  - One or several functions that load data into a map of such struct(s) must be implemented.

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
