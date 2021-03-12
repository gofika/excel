package excel

import (
	. "gopkg.in/check.v1"
)

func (s *XlsxSuite) TestNewFile(c *C) {
	f := NewFile()
	err := f.SaveFile("test_docs/empty.xlsx")
	c.Assert(err, IsNil)
}

func (s *XlsxSuite) TestOpenFile(c *C) {
	f, err := OpenFile("test_docs/two_sheet.xlsx")
	c.Assert(err, IsNil)
	c.Assert(f.xFile.Worksheets, HasLen, 2)
}
