package packaging

import (
	"io/ioutil"
	"path"

	. "gopkg.in/check.v1"
)

const defaultWorkbookRelationshipsTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"></Relationship><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"></Relationship><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"></Relationship></Relationships>`

func testNewWorkbookXRelationships(c *C, file *XFile) {
	var result string
	var err error
	if needWriteTestFile {
		result, err = XMLMarshalAppendHeadIndent(file.WorkbookRelationships)
	} else {
		result, err = XMLMarshalAppendHead(file.WorkbookRelationships)
	}
	c.Assert(err, IsNil)
	if needWriteTestFile {
		err = ioutil.WriteFile(path.Join(templatePath, WorkbookRelationshipsPath, WorkbookRelationshipsFileName), []byte(result), 0644)
		c.Assert(err, IsNil)
	}
	if isAssertDefaultTemplate {
		c.Assert(result, Equals, defaultWorkbookRelationshipsTemplate)
	}
}

const defaultRootRelationshipsTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"></Relationship><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"></Relationship><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"></Relationship></Relationships>`

func testNewDefaultRootXRelationships(c *C, file *XFile) {
	var result string
	var err error
	if needWriteTestFile {
		result, err = XMLMarshalAppendHeadIndent(file.RootRelationships)
	} else {
		result, err = XMLMarshalAppendHead(file.RootRelationships)
	}
	c.Assert(err, IsNil)
	if needWriteTestFile {
		err = ioutil.WriteFile(path.Join(templatePath, RootRelationshipsPath, RootRelationshipsFileName), []byte(result), 0644)
		c.Assert(err, IsNil)
	}
	if isAssertDefaultTemplate {
		c.Assert(result, Equals, defaultRootRelationshipsTemplate)
	}
}
