package xlsrpt_test

import (
	"database/sql"

)

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
