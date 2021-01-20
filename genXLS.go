// Package xlsrpt allows easy generation of Excel Reports.
// Report is generated from a *sql.DB datasource.
package xlsrpt

import (
	"database/sql"
	"errors"
	"fmt"
	"log"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"time"

	"github.com/tealeg/xlsx"
)

// RepColumns - Report Columns Definition.
type RepColumns struct {
	Title   string
	SumFlag bool
}

// RepParams - Parameters for Report Generation.
type RepParams struct {
	RepTitle   string
	RepSheet   string
	RepCols    []RepColumns
	Query      string
	FilePath   string
	AltBg      bool
	AutoFilter bool
	NoTitleRow bool
}

// MultiSheetRep type is used for multiple sheets reports.
type MultiSheetRep struct {
	Params RepParams
	Data   ReportData
	DB     *sql.DB
}

/*
ReportData defines LoadRows() function that must be implemented

Implement this by creating a struct with the fields you want to use as columns in the report.
Then create a map with items of the structure type (map key can be any type).

Note that this is only needed when using ExcelReport() or ExcelMultiSheet() functions.
*/
type ReportData interface {
	LoadRows(rows *sql.Rows) error
}

// Vervose can be used to print information about Excel File generation when set to <true>.
var Vervose bool

// Debug can be used to print debug information about Excel File generation when set to <true>.
var Debug bool

// ExcelReport generates excel report using a datamap that should be loaded by your implementation of LoadRows() function
func ExcelReport(rp RepParams, rptData ReportData, db *sql.DB) error {
	var file *xlsx.File

	file = xlsx.NewFile()

	rows, err := db.Query(rp.Query)
	if err != nil {
		return err
	}

	rptData.LoadRows(rows)
	rows.Close()

	if rp.FilePath == "" {
		rp.FilePath = rp.RepTitle + ".xlsx"
	} else {
		match, _ := regexp.MatchString(`(?m)([a-zA-Z0-9\s_\\.\-\(\):])+(.xls|.xlsx)$`, rp.FilePath)
		if !match {
			rp.FilePath = rp.FilePath + ".xlsx"
		}

		match, _ = regexp.MatchString(`xls$`, rp.FilePath)
		if match {
			fmt.Printf("Warning: File \"%s\" has extension .xls, should be .xlsx\n", rp.FilePath)
		}
	}

	if rp.RepSheet == "" {
		if len(rp.RepTitle) > 30 {
			rp.RepSheet = rp.RepTitle[:30]
		} else {
			rp.RepSheet = rp.RepTitle
		}
	}

	start := time.Now()
	err = genSheet(file, rp, rptData)
	if err != nil {
		fmt.Println("Warning, error generating Excel Sheet:", err.Error())
	}
	if Debug {
		fmt.Println("genSheet() Took:", time.Since(start))
	}

	err = file.Save(rp.FilePath)
	if err != nil {
		return err
	}

	return nil
}

// ExcelMultiSheet generates a Report with Multiple Sheets using a datamap that should be loaded by your implementation of LoadRows() function.
func ExcelMultiSheet(filePath string, reports []MultiSheetRep) error {

	if filePath == "" {
		return errors.New("filePath is empty string")
	}

	var file *xlsx.File

	file = xlsx.NewFile()

	for _, k := range reports {
		rows, err := k.DB.Query(k.Params.Query)
		if err != nil {
			return err
		}

		k.Data.LoadRows(rows)
		rows.Close()

		if k.Params.RepSheet == "" {
			if len(k.Params.RepTitle) > 30 {
				k.Params.RepSheet = k.Params.RepTitle[:30]
			} else {
				k.Params.RepSheet = k.Params.RepTitle
			}
		}

		err = genSheet(file, k.Params, k.Data)
		if err != nil {
			fmt.Println("Warning, error generating Excel Sheet:", err.Error())
		}
	}

	match, _ := regexp.MatchString(`(?m)([a-zA-Z0-9\s_\\.\-\(\):])+(.xls|.xlsx)$`, filePath)
	if !match {
		filePath = filePath + ".xlsx"
	}

	match, _ = regexp.MatchString(`xls$`, filePath)
	if match {
		fmt.Printf("Warning: File \"%s\" has extension .xls, should be .xlsx\n", filePath)
	}

	err := file.Save(filePath)
	if err != nil {
		return err
	}

	return nil
}

/*
ExcelFromDB can be used when the selected columns are not known.
Uses reflect to infer data type directly from DB.
*/
func ExcelFromDB(rp RepParams, db *sql.DB) error {

	driverType := reflect.TypeOf(db.Driver())
	switch driverType.String() {
	case "*mysql.MySQLDriver":
		fmt.Printf("Warning: MySQL Driver is not reflect friendly.\nPlease use ExcelReport() function for MySQL databases.\n")
	}

	var file *xlsx.File

	file = xlsx.NewFile()

	if rp.FilePath == "" {
		rp.FilePath = rp.RepTitle + ".xlsx"
	} else {
		match, _ := regexp.MatchString(`(?m)([a-zA-Z0-9\s_\\.\-\(\):])+(.xls|.xlsx)$`, rp.FilePath)
		if !match {
			rp.FilePath = rp.FilePath + ".xlsx"
		}

		match, _ = regexp.MatchString(`xls$`, rp.FilePath)
		if match {
			fmt.Printf("Warning: File \"%s\" has extension .xls, should be .xlsx\n", rp.FilePath)
		}

	}

	if rp.RepSheet == "" {
		if len(rp.RepTitle) > 30 {
			rp.RepSheet = rp.RepTitle[:30]
		} else {
			rp.RepSheet = rp.RepTitle
		}
	}

	err := genSheetFromDB(file, rp, db)
	if err != nil {
		fmt.Println("Warning, error generating Excel Sheet:", err.Error())
	}

	err = file.Save(rp.FilePath)
	if err != nil {
		return err
	}

	return nil
}

// ExcelMultiSheetFromDB generates a Report with Multiple Sheets.
// Uses reflect to infer data type directly from DB.
func ExcelMultiSheetFromDB(filePath string, reports []MultiSheetRep) error {

	// Make sure they are sorted
	sort.Strings(UntouchCols)
	var file *xlsx.File

	file = xlsx.NewFile()

	for _, k := range reports {
		if k.Params.RepSheet == "" {
			if len(k.Params.RepTitle) > 30 {
				k.Params.RepSheet = k.Params.RepTitle[:30]
			} else {
				k.Params.RepSheet = k.Params.RepTitle
			}
		}
		if Vervose {
			fmt.Println("Adding Sheet", k.Params.RepSheet)
		}

		// TODO(moisoto): Test other drivers (tested on mssql and oracle)
		driverType := reflect.TypeOf(k.DB.Driver())
		switch driverType.String() {
		case "*mysql.MySQLDriver":
			fmt.Printf("On Report Sheet \"%s\" - Warning: MySQL Driver is not reflect friendly.\nPlease use ExcelMultiSheet() function for MySQL databases.\n", k.Params.RepSheet)
		}
		genSheetFromDB(file, k.Params, k.DB)
	}

	match, _ := regexp.MatchString(`(?m)([a-zA-Z0-9\s_\\.\-\(\):])+(.xls|.xlsx)$`, filePath)
	if !match {
		filePath = filePath + ".xlsx"
	}

	match, _ = regexp.MatchString(`xls$`, filePath)
	if match {
		fmt.Printf("Warning: File \"%s\" has extension .xls, should be .xlsx\n", filePath)
	}

	err := file.Save(filePath)
	if err != nil {
		return err
	}

	return nil
}

// genSheet adds the report in a new sheet.
func genSheet(file *xlsx.File, rp RepParams, dataMap interface{}) error {
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var rdata = reflect.ValueOf(dataMap)
	var timeKind = reflect.TypeOf(time.Time{}).Kind()

	if rdata.Kind() != reflect.Map {
		return errors.New("dataMap is not a map")
	}
	sheet, err := file.AddSheet(rp.RepSheet)
	if err != nil {
		return err
	}

	startRow := 1
	if !rp.NoTitleRow {
		startRow = 4
		// Add Report Title
		sheet.AddRow() // Skip a Row
		cell := sheet.AddRow().AddCell()
		s := cell.GetStyle()
		s.Font.Size = 18
		s.Font.Bold = true
		s.ApplyFont = true
		cell.Value = rp.RepTitle
		row = sheet.AddRow()
	}

	// Add Column Titles
	row = sheet.AddRow()
	for _, k := range rp.RepCols {
		cell := row.AddCell()
		s := cell.GetStyle()
		s.Fill.PatternType = "solid"
		s.Fill.FgColor = "004472C4"
		s.Font.Color = "00FFFFFF"
		s.Font.Bold = true
		s.ApplyFill = true
		s.ApplyFont = true
		cell.Value = k.Title
	}

	flag := false

	/*
		// Add Rows (Ordered Rows, fast)
		// This code assumes key is int, is incremental and begins at 1
		// Assures 100% that the report will be generated in the intended order
		var i int
		nitems := rdata.Len()
		for i = 1; i < nitems; i++ {
			v := rdata.MapIndex(reflect.ValueOf(i))
			fmt.Printf("Value: %+v \n", v.Interface())

			if rp.AltBg {
				flag = i%2 == 0
			}
			row = sheet.AddRow()
			addRow(v.Interface(), row, flag)
		}


		// Add Rows (Unordered Rows, flexible, faster)
		// This code is more flexible, works regardless of key used.
		// Order of map items on go is unspecified, so the excel will have unordered rows
		iter := rdata.MapRange()
		for i = 0; iter.Next(); i++ {
			v := iter.Value()
			if rp.AltBg {
				flag = i%2 == 0
			}
			row = sheet.AddRow()
			addRow(v.Interface(), row, flag)
		}
	*/

	// Add Rows (Ordered Rows, flexible, slower)
	// This third implementation intends to accomplish map ordering indepent of key type
	// Implementer of LoadRows can use any column data for map Type, yielding rows ordered by that column
	tkeys := reflect.TypeOf(dataMap).Key() // Type of the Map Key
	rkeys := rdata.MapKeys()               // Slice with Map Keys
	qkeys := len(rkeys)                    // Alternative: rdata.Len()

	//switch <- Kind of Map Keys
	switch tkeys.Kind() {
	case reflect.Int:
		ordKeys := make([]int, qkeys, qkeys)
		for i, k := range rkeys {
			ordKeys[i] = int(k.Int())
		}
		sort.Ints(ordKeys)
		for i, k := range ordKeys {
			v := rdata.MapIndex(reflect.ValueOf(k))

			if rp.AltBg {
				flag = i%2 == 0
			}
			row = sheet.AddRow()
			addRow(v.Interface(), row, flag)
		}
	case reflect.Float32:
		ordKeys := make([]float32, qkeys, qkeys)
		for i, k := range rkeys {
			ordKeys[i] = float32(k.Float())
		}
		sort.Slice(ordKeys, func(i, j int) bool {
			return ordKeys[i] < ordKeys[j]
		})
		for i, k := range ordKeys {
			v := rdata.MapIndex(reflect.ValueOf(k))

			if rp.AltBg {
				flag = i%2 == 0
			}
			row = sheet.AddRow()
			addRow(v.Interface(), row, flag)
		}
	case reflect.Float64:
		ordKeys := make([]float64, qkeys, qkeys)
		for i, k := range rkeys {
			ordKeys[i] = float64(k.Float())
		}
		sort.Float64s(ordKeys)
		for i, k := range ordKeys {
			v := rdata.MapIndex(reflect.ValueOf(k))

			if rp.AltBg {
				flag = i%2 == 0
			}
			row = sheet.AddRow()
			addRow(v.Interface(), row, flag)
		}
	case reflect.String:
		ordKeys := make([]string, qkeys, qkeys)
		for i, k := range rkeys {
			ordKeys[i] = (k.String())
		}
		sort.Strings(ordKeys)
		for i, k := range ordKeys {
			v := rdata.MapIndex(reflect.ValueOf(k))

			if rp.AltBg {
				flag = i%2 == 0
			}
			row = sheet.AddRow()
			addRow(v.Interface(), row, flag)
		}
	case timeKind:
		ordKeys := make([]time.Time, qkeys, qkeys)
		for i, k := range rkeys {
			kTime := k.Interface().(time.Time)
			ordKeys[i] = time.Time(kTime)
		}
		sort.Slice(ordKeys, func(i, j int) bool {
			return ordKeys[i].Before(ordKeys[j])
		})
		for i, k := range ordKeys {
			v := rdata.MapIndex(reflect.ValueOf(k))

			if rp.AltBg {
				flag = i%2 == 0
			}
			row = sheet.AddRow()
			addRow(v.Interface(), row, flag)
		}
	default:
		return fmt.Errorf("dataMap key not a valid kind (%+v)", reflect.TypeOf(dataMap).Key().Kind())
	}

	_ = sheet.SetColWidth(0, len(rp.RepCols)-1, 28.0)
	if rp.AutoFilter {
		var brCell string
		c := len(rp.RepCols)
		for i := 65; c > 26; c -= 26 {
			brCell = string(i)
			i++
		}
		brCell = brCell + string(64+c) + strconv.Itoa(qkeys+startRow)
		tpCell := "A" + strconv.Itoa(startRow)
		sheet.AutoFilter = &xlsx.AutoFilter{TopLeftCell: tpCell, BottomRightCell: brCell}
	}
	/*
		if rp.AutoFilter {
			brCell := string(64+len(rp.RepCols)) + strconv.Itoa(qkeys+4)
			sheet.AutoFilter = &xlsx.AutoFilter{TopLeftCell: "A4", BottomRightCell: brCell}
		}
	*/
	if qkeys != 0 { // If there's Data to be Processed
		row = sheet.AddRow()

		for c, col := range rp.RepCols {
			colLetter := string(c + 65)
			cell := row.AddCell()
			s := cell.GetStyle()
			s.Fill.PatternType = "solid"
			s.Fill.FgColor = "00D0CECE"
			s.ApplyFill = true
			if col.SumFlag {
				formula := "=SUBTOTAL(109," + colLetter + strconv.Itoa(startRow+1) + ":" + colLetter + strconv.Itoa(startRow+qkeys) + ")"
				cell.SetFloatWithFormat(0, "$#,##0.00")
				cell.SetFormula(formula)
				s.Font.Bold = true
				s.Font.Color = "00FF0000"
				s.Alignment.Horizontal = "left"
				s.ApplyAlignment = true
				s.ApplyFont = true
			}
		}
	}
	return nil
}

func genSheetFromDB(file *xlsx.File, rp RepParams, db *sql.DB) error {
	var sheet *xlsx.Sheet
	var row *xlsx.Row

	sheet, err := file.AddSheet(rp.RepSheet)
	if err != nil {
		return err
	}

	rows, err := db.Query(rp.Query)
	if err != nil {
		return err
	}

	cols, err := rows.Columns()
	if err != nil {
		return err
	}

	startRow := 1
	if !rp.NoTitleRow {
		startRow = 4
		// Add Report Title
		sheet.AddRow() // Skip a Row
		cell := sheet.AddRow().AddCell()
		s := cell.GetStyle()
		s.Font.Size = 18
		s.Font.Bold = true
		s.ApplyFont = true
		cell.Value = rp.RepTitle
		row = sheet.AddRow()
	}

	// Add Titles
	row = sheet.AddRow()
	for _, k := range cols {
		cell := row.AddCell()
		s := cell.GetStyle()
		s.Fill.PatternType = "solid"
		s.Fill.FgColor = "004472C4"
		s.Font.Color = "00FFFFFF"
		s.Font.Bold = true
		s.ApplyFill = true
		s.ApplyFont = true
		cell.Value = k
	}

	var i int
	flag := false
	for rows.Next() {
		// Create a slice of interface{}'s to represent each column,
		// and a second slice to contain pointers to each item in the columns slice.
		columns := make([]interface{}, len(cols))
		columnPointers := make([]interface{}, len(cols))
		for i := range columns {
			columnPointers[i] = &columns[i]
		}
		// Scan the result into the column pointers...
		if err = rows.Scan(columnPointers...); err != nil {
			return err
		}

		// Create our map, and retrieve the value for each column from the pointers slice,
		// storing it in the map with the name of the column as the key.
		m := make(map[string]interface{})
		for i, colName := range cols {
			val := columnPointers[i].(*interface{})
			m[colName] = *val
		}

		if rp.AltBg {
			flag = i%2 == 0
		}
		row = sheet.AddRow()
		addMapRow(cols, m, row, flag) // v.Interface(), row, flag)
		i++
		//fmt.Println("Processing Line:", i)
	}

	//fmt.Println("Report Lines Quantity:", i)
	_ = sheet.SetColWidth(0, len(cols)-1, 28.0)
	if rp.AutoFilter {
		var brCell string
		c := len(cols)
		for i = 65; c > 26; c -= 26 {
			brCell = string(i)
			i++
		}
		brCell = brCell + string(64+c) + strconv.Itoa(i+startRow)
		tpCell := "A" + strconv.Itoa(startRow)
		sheet.AutoFilter = &xlsx.AutoFilter{TopLeftCell: tpCell, BottomRightCell: brCell}
	}

	if i != 0 { // If there's Data to be Processed
		row = sheet.AddRow()

		for c := 0; c < len(cols); c++ {
			//for c, col := range cols {
			//colLetter := string(c + 65)
			cell := row.AddCell()
			s := cell.GetStyle()
			s.Fill.PatternType = "solid"
			s.Fill.FgColor = "00D0CECE"
			s.ApplyFill = true
			/*
				if col.SumFlag {
					formula := "SUM(" + colLetter + "2:" + colLetter + strconv.Itoa(i+1) + ")"
					cell.SetFloatWithFormat(0, "$#,##0.00")
					cell.SetFormula(formula)
					s.Font.Bold = true
					s.Font.Color = "00FF0000"
					s.Alignment.Horizontal = "left"
					s.ApplyAlignment = true
					s.ApplyFont = true
				}
			*/
		}
	}

	return nil
}

func runningtime(s string) (string, time.Time) {
	log.Println("Start:	", s)
	return s, time.Now()
}

func track(s string, startTime time.Time) {
	endTime := time.Now()
	log.Println("End:	", s, "took", endTime.Sub(startTime))
}
