package excel

import (
	. "gopkg.in/check.v1"
)

func (s *XlsxSuite) TestColumnTitle(c *C) {
	c.Assert(ColumnName(0), Equals, "")
	c.Assert(ColumnName(1), Equals, "A")
	c.Assert(ColumnName(26), Equals, "Z")
	c.Assert(ColumnName(26*2+1), Equals, "BA")
	c.Assert(ColumnName(26*3+1), Equals, "CA")
	c.Assert(ColumnName(26*26+1), Equals, "ZA")
}

func (s *XlsxSuite) TestColumnNumber(c *C) {
	c.Assert(ColumnNumber("WrongNumber"), Equals, 0)
	c.Assert(ColumnNumber("A"), Equals, 1)
	c.Assert(ColumnNumber("Z"), Equals, 26)
	c.Assert(ColumnNumber("AA"), Equals, 26+1)
	c.Assert(ColumnNumber("BA"), Equals, 26*2+1)
	c.Assert(ColumnNumber("CA"), Equals, 26*3+1)
	c.Assert(ColumnNumber("ZA"), Equals, 26*26+1)
}

func (s *XlsxSuite) TestCellNameToCoordinates(c *C) {
	var col, row int
	col, row = CellNameToCoordinates("A5")
	c.Assert(col, Equals, 1)
	c.Assert(row, Equals, 5)

	col, row = CellNameToCoordinates("Z9")
	c.Assert(col, Equals, 26)
	c.Assert(row, Equals, 9)
}
