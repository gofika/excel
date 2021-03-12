package packaging

import (
	"io/ioutil"
	"path"

	. "gopkg.in/check.v1"
)

const defaultWorkbookTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"></fileVersion><workbookPr defaultThemeVersion="164011"></workbookPr><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="22260" windowHeight="12645"></workbookView></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"></sheet></sheets><calcPr calcId="162913"></calcPr></workbook>`

func testNewXWorkbook(c *C, file *XFile) {
	var result string
	var err error
	if needWriteTestFile {
		result, err = XMLMarshalAppendHeadIndent(file.Workbook)
	} else {
		result, err = XMLMarshalAppendHead(file.Workbook)
	}
	c.Assert(err, IsNil)
	if needWriteTestFile {
		err = ioutil.WriteFile(path.Join(templatePath, WorkbookPath, WorkbookFileName), []byte(result), 0644)
		c.Assert(err, IsNil)
	}
	if isAssertDefaultTemplate {
		c.Assert(result, Equals, defaultWorkbookTemplate)
	}
}
