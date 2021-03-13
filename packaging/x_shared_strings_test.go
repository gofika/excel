package packaging

import (
	. "gopkg.in/check.v1"
	"io/ioutil"
	"path"
)

const defaultSharedStringsTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>`

func testNewXSharedStrings(c *C, file *XFile) {
	var result string
	var err error
	result, err = XMLMarshalAppendHeadIndent(file.SharedStrings)
	c.Assert(err, IsNil)
	if needWriteTestFile {
		err = ioutil.WriteFile(path.Join(templatePath, SharedStringsPath, SharedStringsFileName), []byte(result), 0644)
		c.Assert(err, IsNil)
	}
	c.Assert(result, Equals, defaultSharedStringsTemplate)
}
