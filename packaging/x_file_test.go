package packaging

import (
	. "gopkg.in/check.v1"
)

func (s *PackagingSuite) TestFile(c *C) {
	file := NewDefaultFile()
	testTheme(c, file)
	testNewXContentTypes(c, file)
	testNewDefaultXCoreProperties(c, file)
	testNewXExtendedProperties(c, file)
	testNewWorkbookXRelationships(c, file)
	testNewDefaultRootXRelationships(c, file)
	testNewXWorkbook(c, file)
	testNewDefaultXWorksheet(c, file)
	testNewDefaultXStyleSheet(c, file)
	testNewXSharedStrings(c, file)
}
