package excel

import (
	"fmt"
	"strconv"
	"time"

	"github.com/leaker/excel/packaging"
)


// cellImpl cell operator
type cellImpl struct {
	sheet    *sheetImpl
	col      int
	row      int
	cellName string
}

func newCell(sheet *sheetImpl, col, row int) *cellImpl {
	return &cellImpl{
		sheet:    sheet,
		col:      col,
		row:      row,
		cellName: CoordinatesToCellName(col, row),
	}
}

// Row cell row number
func (c *cellImpl) Row() int {
	return c.row
}

// Col cell col number
func (c *cellImpl) Col() int {
	return c.col
}

func (c *cellImpl) getRow() *packaging.XRow {
	return c.sheet.getRow(c.row)
}

func (c *cellImpl) getCell() *packaging.XC {
	return c.sheet.getCell(c.col, c.row)
}

func (c *cellImpl) getSharedStrings() *sharedStrings {
	return newSharedStrings(c.sheet.file)
}

func (c *cellImpl) prepareCell() *packaging.XC {
	return c.sheet.prepareCell(c.col, c.row)
}

// SetValue provides to set the value of a cell
// Allow Types:
//     int
//     int8
//     int16
//     int32
//     int64
//     uint
//     uint8
//     uint16
//     uint32
//     uint64
//     float32
//     float64
//     string
//     []byte
//     time.Duration
//     time.Time
//     bool
//     nil
//
// Example:
//     cell.SetValue(100)
//     cell.SetValue("Hello")
//     cell.SetValue(3.14)
func (c *cellImpl) SetValue(value interface{}) {
	switch v := value.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		c.setIntType(v)
	case float32:
		c.SetFloatValuePrec(float64(v), -1, 32)
	case float64:
		c.SetFloatValue(v)
	case string:
		c.SetStringValue(v)
	case []byte:
		c.SetStringValue(string(v))
	case time.Duration:
		c.SetDurationValue(v)
	case time.Time:
		c.SetTimeValue(v)
	case bool:
		c.SetBoolValue(v)
	case nil:
		c.SetDefaultValue("")
	default:
		c.SetStringValue(fmt.Sprint(value))
	}
	return
}

func (c *cellImpl) setIntType(value interface{}) {
	switch v := value.(type) {
	case int:
		c.SetIntValue(v)
	case int8:
		c.SetIntValue(int(v))
	case int16:
		c.SetIntValue(int(v))
	case int32:
		c.SetIntValue(int(v))
	case int64:
		c.SetIntValue(int(v))
	case uint:
		c.SetIntValue(int(v))
	case uint8:
		c.SetIntValue(int(v))
	case uint16:
		c.SetIntValue(int(v))
	case uint32:
		c.SetIntValue(int(v))
	case uint64:
		c.SetIntValue(int(v))
	}
}

// SetIntValue set cell for int type
func (c *cellImpl) SetIntValue(value int) {
	cell := c.prepareCell()
	cell.T = ""
	cell.V = strconv.Itoa(value)
}

// GetIntValue get cell value with int type
func (c *cellImpl) GetIntValue() int {
	cell := c.getCell()
	if cell == nil {
		return 0
	}
	value, err := strconv.Atoi(cell.V)
	if err != nil {
		return 0
	}
	return value
}

// SetFloatValue set cell for float64 type
func (c *cellImpl) SetFloatValue(value float64) {
	c.SetFloatValuePrec(value, -1, 64)
}

// SetFloatValuePrec set cell for float64 type with pres
func (c *cellImpl) SetFloatValuePrec(value float64, prec int, bitSize int) {
	cell := c.prepareCell()
	cell.V = strconv.FormatFloat(value, 'f', prec, bitSize)
}

// GetStringValue get cell value with string type
func (c *cellImpl) GetStringValue() string {
	cell := c.getCell()
	if cell == nil {
		return ""
	}
	return c.getSharedStrings().Get(cell.V)
}

// SetStringValue set cell value for string type
func (c *cellImpl) SetStringValue(value string) {
	cell := c.prepareCell()
	cell.T = "s"
	stringID := c.getSharedStrings().Append(value)
	cell.V = stringID
}

// SetBoolValue set cell value for bool type
func (c *cellImpl) SetBoolValue(value bool) {
	cell := c.prepareCell()
	cell.T = "b"
	if value {
		cell.V = "1"
	} else {
		cell.V = "0"
	}
}

// SetDefaultValue set cell value without any type
func (c *cellImpl) SetDefaultValue(value string) {
	cell := c.prepareCell()
	cell.V = value
}

// SetTimeValue set cell value for time.Time type
func (c *cellImpl) SetTimeValue(value time.Time) {
	cell := c.prepareCell()
	cell.T = ""

	excelTime := TimeToExcelTime(value)
	if excelTime > 0 {
		cell.V = strconv.FormatFloat(excelTime, 'f', -1, 64)
	} else {
		cell.V = value.Format(time.RFC3339Nano)
	}
}

// SetDurationValue set cell value for time.Duration type
func (c *cellImpl) SetDurationValue(value time.Duration) {
	cell := c.prepareCell()
	cell.V = strconv.FormatFloat(value.Seconds()/86400.0, 'f', -1, 32)
	// TODO: update cell style
	// c.setDefaultTimeStyle(21)
}
