package xlsrpt_test

import (
	"github.com/moisoto/xlsrpt"
)

func ExampleExcelFromDB() {
	repParams := xlsrpt.RepParams{
		RepTitle:   "Customer Report",
		Query:      "SELECT CreationDate, FirstName, LastName, CustomerNumber, Balance FROM Customer;",
		NoTitleRow: true,
		AutoFilter: true}

	// Open Connection to Database
	database, err := dbConnect()
	if err != nil {
		panic(err.Error())
	}

	// Just call ExcelFromDB and pass the parameters and the *sql.DB pointer
	xlsrpt.ExcelFromDB(repParams, database)
}

/*
func dbConnect() (*sql.DB, error) {
	var db *sql.DB
	driverName := "mssql" // also tested with oracle driver
	dataSourceName := "your connection string"

	db, err := sql.Open(driverName, dataSourceName)
	// if there is an error opening the connection, handle it
	if err != nil {
		return nil, err
	}

	// defer the close till after the main function has finished
	// executing
	defer db.Close()
	return db, nil
}
*/
