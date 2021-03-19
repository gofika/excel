package excel

import (
	. "gopkg.in/check.v1"
	"time"
)

func (s *XlsxSuite) TestNewSheet(c *C) {
	const docPath = "test_docs/two_sheet.xlsx"

	f := NewFile()
	f.NewSheet("Sheet2")
	err := f.SaveFile(docPath)
	c.Assert(err, IsNil)

	f, err = OpenFile(docPath)
	c.Assert(err, IsNil)
	sheet := f.OpenSheet("Sheet2")
	c.Assert(sheet, NotNil)
}

func (s *XlsxSuite) TestSetCellValue(c *C) {
	const docPath = "test_docs/set_cell_values.xlsx"
	f := NewFile()

	sheet := f.OpenSheet("Sheet1")
	sheet.SetCellValue(ColumnNumber("A"), 1, "Name")
	sheet.SetCellValue(ColumnNumber("B"), 1, "Score")
	sheet.SetCellValue(ColumnNumber("A"), 2, "Jason")
	sheet.SetCellValue(ColumnNumber("B"), 2, 100)

	sheet.SetCellValue(ColumnNumber("C"), 3, 200.50)
	sheet.SetCellValue(ColumnNumber("D"), 3, time.Date(1980, 9, 8, 23, 40, 10, 40, time.Local))
	sheet.SetCellValue(ColumnNumber("E"), 4, 10*time.Second)

	cellD4 := sheet.Cell(ColumnNumber("D"), 4)
	cellD4.SetTimeValue(time.Date(1980, 9, 8, 23, 40, 10, 40, time.Local))
	cellD4.SetNumberFormat("yyyy-mm-dd hh:mm:ss")
	err := f.SaveFile(docPath)
	c.Assert(err, IsNil)
}
