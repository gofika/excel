package excel

import (
	. "gopkg.in/check.v1"
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
