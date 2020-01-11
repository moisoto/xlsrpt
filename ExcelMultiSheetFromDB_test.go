package xlsrpt_test

import (
	"github.com/moisoto/xlsrpt"
)

func ExampleExcelMultiSheetFromDB() {
	// Open Connection to Database
	database, err := dbConnect()
	if err != nil {
		panic(err.Error())
	}

	repParams := []xlsrpt.MultiSheetRep{
		{
			Params: xlsrpt.RepParams{
				RepTitle:   "All Customers",
				Query:      "SELECT CreationDate, FirstName, LastName, CustomerNumber, Balance FROM Customer;",
				NoTitleRow: true,
				AutoFilter: true},
			DB: database},
		{
			Params: xlsrpt.RepParams{
				RepTitle:   "VIP Customers",
				Query:      "SELECT CreationDate, FirstName, LastName, CustomerNumber, Balance FROM Customer WHERE vip=1;",
				NoTitleRow: false,
				AutoFilter: true},
			DB: database}}

	// Just call ExcelMultiSheetFromDB and pass the file name and report parameters.
	xlsrpt.ExcelMultiSheetFromDB("Customer Report.xlsx", repParams)
}
