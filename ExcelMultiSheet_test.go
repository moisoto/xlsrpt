package xlsrpt_test

import (
	"github.com/moisoto/xlsrpt"
)

func ExampleExcelMultiSheet() {

	// Open Connection to Database
	database, err := dbConnect()
	if err != nil {
		panic(err.Error())
	}

	// repExampleMap and LoadRows() implementation for repExampleMap is located on ReportData_test.go
	// Here we use the same struct (repExampleMap) for simplification since example columns are the same on both sheets
	// Normally you will have several structs and several implementations of LoadRows() for each struct
	var rptDataMap1 = make(repExampleMap)
	var rptDataMap2 = make(repExampleMap)

	var rptData1 xlsrpt.ReportData = rptDataMap1
	var rptData2 xlsrpt.ReportData = rptDataMap2

	repParams := []xlsrpt.MultiSheetRep{
		{
			Params: xlsrpt.RepParams{
				RepTitle: "All Accounts",
				RepCols: []xlsrpt.RepColumns{
					{Title: "Date Created", SumFlag: false},
					{Title: "First Name", SumFlag: false},
					{Title: "Last Name", SumFlag: false},
					{Title: "Customer Number", SumFlag: false},
					{Title: "Customer Balance", SumFlag: true}},
				Query: "SELECT CreationDate, FirstName, LastName, CustomerNumber, Balance FROM Customer;",
				AutoFilter: true},
			Data: rptData1,
			DB:   database},
		{
			Params: xlsrpt.RepParams{
				RepTitle: "VIP Accounts",
				RepSheet: "VIPs", // You can specify a sheet name, otherwise RepTitle will be used as the sheet name
				RepCols: []xlsrpt.RepColumns{
					{Title: "Date Created", SumFlag: false},
					{Title: "First Name", SumFlag: false},
					{Title: "Last Name", SumFlag: false},
					{Title: "Customer Number", SumFlag: false},
					{Title: "Customer Balance", SumFlag: true}},
				Query: "SELECT CreationDate, FirstName, LastName, CustomerNumber, Balance FROM Customer WHERE vip=1;",
				AutoFilter: true},
			Data: rptData2,
			DB:   database}}

	// Just call ExcelMultiSheet and pass the file name and report parameters. 
	xlsrpt.ExcelMultiSheet("Customer Report.xlsx", repParams)
}