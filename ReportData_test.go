package xlsrpt_test

import (
	"database/sql"
	"fmt"
	"strconv"

	"github.com/moisoto/xlsrpt"
)

type repExampleData struct {
	DateCreated   xlsrpt.CellStr
	FirstName     xlsrpt.CellStr
	LastName      xlsrpt.CellStr
	AccountNumber xlsrpt.CellInt
	Balance       xlsrpt.CellCurrency
}

type repExampleMap map[string]repExampleData

func ExampleReportData() {
	fmt.Println("Example of Implementation of ReportData interface.")
	// Output: Example of Implementation of ReportData interface.
}

// LoadRows implements LoadRows function using type repExampleMap
func (dataMap repExampleMap) LoadRows(rows *sql.Rows) error {
	mapIndex := 0
	for rows.Next() {
		mapIndex++
		var d repExampleData
		err := rows.Scan(
			&d.DateCreated, &d.FirstName, &d.LastName, &d.AccountNumber, &d.Balance)
		if err != nil {
			fmt.Println(err.Error())
			rows.Close()
			return nil
		}
		dataMap[strconv.Itoa(mapIndex)] = d
	}
	return nil
}
