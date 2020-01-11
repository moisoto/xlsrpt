package xlsrpt_test

import (
	"github.com/moisoto/xlsrpt"
)

func ExampleExcelReport() {
	repParams := xlsrpt.RepParams{
		RepTitle: "Customer Report",
		RepCols: []xlsrpt.RepColumns{
			{Title: "Date Created", SumFlag: false},
			{Title: "First Name", SumFlag: false},
			{Title: "Last Name", SumFlag: false},
			{Title: "Customer Number", SumFlag: false},
			{Title: "Customer Balance", SumFlag: true}},
		Query: "SELECT CreationDate, FirstName, LastName, CustomerNumber, Balance FROM Customer;"}

	// Open Connection to Database
	database, err := dbConnect()
	if err != nil {
		panic(err.Error())
	}

	// repExampleMap and LoadRows() implementation for repExampleMap is located on ReportData_test.go
	var rptDataMap = make(repExampleMap)
	var rptData xlsrpt.ReportData = rptDataMap

	// Call ExcelReport with the report parameters, dataMap and database pointer.
	// Check ReportData_test.go to see how to implement the LoadRows() function.
	xlsrpt.ExcelReport(repParams, rptData, database)
}
