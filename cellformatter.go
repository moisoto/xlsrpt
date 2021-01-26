package xlsrpt

import (
	"fmt"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

// CellInt - Integer Cell Type.
type CellInt int

// CellStr - String Cell Type.
type CellStr string

// CellNumeric - Numeric Cell Type (Number with no format).
type CellNumeric float64

// CellDecimal - Decimal Cell Type (Number with commas and 2 decimal places).
type CellDecimal float64

// CellPercent - Percent Cell Type.
type CellPercent float64

// CellCurrency - Currency Cell Type.
type CellCurrency float32

// CellDate - Date Cell Type.
type CellDate time.Time

/*
// *** No Need for this, since we are not using cellAdder objects (yet) ***
// *** Will leave commented just in case ***
// cellAdder interface allows for easier addition of fields to cell with
// proper formatting implemented on types:
// cellInt, cellStr, cellNumber, cellCurrency, cellDecimal
type cellAdder interface {
	addCell(row *xlsx.Row)
}
*/

// Library behavior configuration variables.
var (
	// LogBench can be used to log benchmark information of report creation (unimplemented).
	LogBench bool

	// UntouchStrings can be used to leave strings untouched.
	UntouchStrings bool

	// UntouchCols can be used to set columns that must not be formatted.
	UntouchCols []string
)

// AddRow adds a row to excel report.
func addRow(fields interface{}, row *xlsx.Row, flag bool) {
	var timeKind = reflect.TypeOf(time.Time{}).Kind()

	if reflect.ValueOf(fields).Kind() != reflect.Struct {
		fmt.Println("Logic Error. It's not a Struct!")
		return
	}

	f := reflect.ValueOf(fields)

	for i := 0; i < f.NumField(); i++ {
		switch f.Field(i).Kind() {
		case reflect.Int:
			v := CellInt(f.Field(i).Int())
			altBgColor(v.addCell(row), flag)
		case reflect.String:
			v := CellStr(f.Field(i).String())
			altBgColor(v.addCell(row), flag)
		case reflect.Float64:
			v := CellDecimal(f.Field(i).Float())
			altBgColor(v.addCell(row), flag)
		case reflect.Float32:
			v := CellCurrency(f.Field(i).Float())
			altBgColor(v.addCell(row), flag)
		case timeKind:
			v := CellDate(f.Field(i).Interface().(CellDate))
			altBgColor(v.addCell(row), flag)
		default:
			v := CellStr("unimplemented")
			altBgColor(v.addCell(row), flag)
		}
	}

}

func addMapRow(ordColumns []string, mapRow map[string]interface{}, row *xlsx.Row, flag bool) {
	var timeKind = reflect.TypeOf(time.Time{}).Kind()

	// Preload length of UntouchCols so we don't do it for each column
	l := len(UntouchCols)

	for _, v := range ordColumns {
		val := reflect.ValueOf(mapRow[v])
		switch val.Kind() {
		case reflect.Int64:
			v := CellInt(val.Int())
			altBgColor(v.addCell(row), flag)
		case reflect.String:
			goStr := true
			untouchCol := false
			str := val.String()
			if l > 0 {
				i := sort.SearchStrings(UntouchCols, v)
				//fmt.Println("Checking untouchCols for Column", v, )
				found := (i < l && UntouchCols[i] == v)
				if found {
					untouchCol = true
				}
			}
			if !UntouchStrings && !untouchCol {
				// fmt.Println("Formatting Column", v, "original value:", sort.SearchStrings(UntouchCols, v))
				goStr = false
				nType, f := isNum(str)
				switch nType {
				case 'i':
					v := CellInt(int(f))
					altBgColor(v.addCell(row), flag)
				case 'd':
					v := CellNumeric(f)
					altBgColor(v.addCell(row), flag)
				case 'p':
					v := CellPercent(f)
					altBgColor(v.addCell(row), flag)
				default:
					goStr = true
				}
				/*
					// Use isNumDot instead ParseFloat here since it's faster. Only use ParseFloat if number is detected
					if isNumDot(str) { // Numeric Strings with decimal places are treated as percent values
						f, err := strconv.ParseFloat(str, 64)
						if err == nil {
							v := CellNumeric(f)
							altBgColor(v.addCell(row), flag)
							goStr = false // ParseFloat didn't fail don't add as String
						}
					}
					isPerc, f := isNumPerc(str)
					if isPerc { // Numeric Strings with decimal places are treated as percent values
						v := CellPercent(f)
						altBgColor(v.addCell(row), flag)
						goStr = false // ParseFloat didn't fail don't add as String
					}
				*/
			}
			if goStr {
				// TODO(moisoto): Debug ineffectual assignment. Check if it was supposed to revert to false
				// goStr = true

				v := CellStr(val.String())
				altBgColor(v.addCell(row), flag)
			}
		case reflect.Float64:
			v := CellCurrency(val.Float())
			altBgColor(v.addCell(row), flag)
		case timeKind:
			v := val.Interface().(time.Time)
			cell := row.AddCell()
			s := cell.GetStyle()
			s.Alignment.Horizontal = "left"
			s.ApplyAlignment = true
			cell.SetDateTime(v)
			altBgColor(cell, flag)
		default:
			if Vervose {
				fmt.Printf("Invalid Column Type \"%v\" for Column \"%s\" of Value \"%v\"\n", val.Kind(), v, val)
			}
			var empty CellStr
			altBgColor(empty.addCell(row), flag)
		}
	}

}

func (data CellInt) addCell(row *xlsx.Row) (cell *xlsx.Cell) {
	cell = row.AddCell()
	s := cell.GetStyle()
	s.Alignment.Horizontal = "left"
	cell.SetInt(int(data))
	return cell
}

func (data CellStr) addCell(row *xlsx.Row) (cell *xlsx.Cell) {
	cell = row.AddCell()
	s := cell.GetStyle()
	s.Alignment.Horizontal = "left"
	cell.Value = string(data)
	return cell
}

func (data CellNumeric) addCell(row *xlsx.Row) (cell *xlsx.Cell) {
	return addFloatCell(float64(data), "", row)
}

func (data CellDecimal) addCell(row *xlsx.Row) (cell *xlsx.Cell) {
	return addFloatCell(float64(data), "#,##0", row)
}

func (data CellPercent) addCell(row *xlsx.Row) (cell *xlsx.Cell) {
	return addFloatCell(float64(data), "0.00%", row)
}

func (data CellCurrency) addCell(row *xlsx.Row) (cell *xlsx.Cell) {
	return addFloatCell(float64(data), "$#,##0.00", row)
}

func (data CellDate) addCell(row *xlsx.Row) (cell *xlsx.Cell) {
	cell = row.AddCell()
	s := cell.GetStyle()
	s.Alignment.Horizontal = "left"
	s.ApplyAlignment = true

	cell.SetDateTime(time.Time(data))
	return cell
}

func addFloatCell(data float64, format string, row *xlsx.Row) (cell *xlsx.Cell) {
	cell = row.AddCell()
	s := cell.GetStyle()
	s.Alignment.Horizontal = "left"
	if format != "" {
		cell.SetFloatWithFormat(data, format)
	} else {
		cell.SetFloat(data)
	}
	return cell
}

func altBgColor(cell *xlsx.Cell, flag bool) {
	var s *xlsx.Style

	if flag == true {
		s = cell.GetStyle()
		s.Fill.PatternType = "solid"
		s.Fill.FgColor = "00B4C6E7"
		s.ApplyFill = true
	}
}

// returns true if string is numeric and constains a dot.
func isNum(s string) (nType int32, val float64) {
	dotFound := false
	percFound := false

	dot := strings.Count(s, ".")
	switch dot {
	case 0:
		// Do Nothing if no dot
	case 1:
		dotFound = true
	default:
		// More than one dot
		return 0, 0
	}

	perc := strings.Index(s, "%")

	if perc == len(s)-1 {
		percFound = true
	} else if perc > -1 {
		return 0, 0
	}

	for _, v := range s {
		if v != '.' && v != '%' && (v<'0' || v >'9') {
			return 0, 0
		}
		/*
		switch v {
		case '.':
		case '%':
		default:
			// If not '.' nor '%' nor
			if v < '0' || v > '9' {
				return 0, 0
			}
		}
		*/
	}

	fmt.Println("To the meat!")

	if dotFound {
		if !percFound {
			nType = 'd'
		} else {
			s = s[0 : len(s)-1]
			nType = 'p'
		}
	} else {
		nType = 'i'
	}

	f, err := strconv.ParseFloat(s, 64)
	if err != nil || (nType == 'p' && f > 1) {
		return 0, 0
	}

	return nType, f
}
