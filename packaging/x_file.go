package packaging

const (
	// XMLHeader XML Document Header
	XMLHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
)

// XFile Document all Data
type XFile struct {
	ContentTypes          *XContentTypes       // [Content_Types].xml
	Worksheets            []*XWorksheet        // xl/worksheets/sheet?.xml
	Workbook              *XWorkbook           // xl/workbook.xml
	WorkbookRelationships *XRelationships      // xl/_rels/workbook.xml.rels
	RootRelationships     *XRelationships      // _rels/.rels
	ExtendedProperties    *XExtendedProperties // docProps/app.xml
	CoreProperties        *XCoreProperties     // docProps/core.xml
	Themes                []*XTheme            // xl/theme/theme?.xml
	StyleSheet            *XStyleSheet         // xl/styles.xml
	SharedStrings         *XSharedStrings      // xl/sharedStrings.xml
}

// NewDefaultFile create *XFile with default template
func NewDefaultFile() (file *XFile) {
	sheet1 := NewDefaultXWorksheet()
	worksheets := []*XWorksheet{sheet1}

	theme1 := NewDefaultXTheme()
	themes := []*XTheme{theme1}

	file = &XFile{
		Worksheets:     worksheets,
		CoreProperties: NewDefaultXCoreProperties(),
		Themes:         themes,
		StyleSheet:     NewDefaultXStyleSheet(),
		SharedStrings:  NewDefaultXSharedStrings(),
	}

	file.WorkbookRelationships = NewWorkbookXRelationships(file)
	Workbook := NewXWorkbook(file.WorkbookRelationships)
	file.Workbook = Workbook
	file.ExtendedProperties = NewXExtendedProperties(Workbook)
	file.ContentTypes = NewXContentTypes(file.WorkbookRelationships)
	file.RootRelationships = NewDefaultRootXRelationships()
	return
}
