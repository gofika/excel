package packaging

import (
	"github.com/gofika/util/fileutil"
	"path"

	. "gopkg.in/check.v1"
)

const defaultExtendedPropertiesTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Microsoft Excel</Application>
    <DocSecurity>0</DocSecurity>
    <ScaleCrop>false</ScaleCrop>
    <HeadingPairs>
        <vt:vector size="2" baseType="variant">
            <vt:variant>
                <vt:lpstr>Worksheets</vt:lpstr>
            </vt:variant>
            <vt:variant>
                <vt:i4>1</vt:i4>
            </vt:variant>
        </vt:vector>
    </HeadingPairs>
    <TitlesOfParts>
        <vt:vector size="1" baseType="lpstr">
            <vt:lpstr>Sheet1</vt:lpstr>
        </vt:vector>
    </TitlesOfParts>
    <Company></Company>
    <LinksUpToDate>false</LinksUpToDate>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>16.0300</AppVersion>
</Properties>`

func testNewXExtendedProperties(c *C, file *XFile) {
	var result string
	var err error
	result, err = XMLMarshalAppendHeadIndent(file.ExtendedProperties)
	c.Assert(err, IsNil)
	if needWriteTestFile {
		err = fileutil.WriteFile(path.Join(templatePath, ExtendedPropertiesPath, ExtendedPropertiesFileName), []byte(result))
		c.Assert(err, IsNil)
	}
	c.Assert(result, Equals, defaultExtendedPropertiesTemplate)
}
