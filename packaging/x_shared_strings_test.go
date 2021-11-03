package packaging

import (
	"github.com/gofika/util/fileutil"
	. "gopkg.in/check.v1"
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
		err = fileutil.WriteFile(path.Join(templatePath, SharedStringsPath, SharedStringsFileName), []byte(result))
		c.Assert(err, IsNil)
	}
	c.Assert(result, Equals, defaultSharedStringsTemplate)
}
