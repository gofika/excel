package packaging

import "encoding/xml"

// StyleSheet Defines
const (
	StyleSheetContentType      = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
	StyleSheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

	StyleSheetPath     = "xl"
	StyleSheetFileName = "styles.xml"
)

// XStyleSheet StyleSheet XML document
type XStyleSheet struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	XmlnsMc     string   `xml:"xmlns:mc,attr"`
	McIgnorable string   `xml:"mc:Ignorable,attr"`
	XmlnsX14ac  string   `xml:"xmlns:x14ac,attr"`
	XmlnsX16r2  string   `xml:"xmlns:x16r2,attr"`

	Fonts        *XFonts            `xml:"fonts"`
	Fills        *XFills            `xml:"fills"`
	Borders      *XBorders          `xml:"borders"`
	CellStyleXfs *XCellStyleXfs     `xml:"cellStyleXfs"`
	CellXfs      *XCellXfs          `xml:"cellXfs"`
	CellStyles   *XCellStyles       `xml:"cellStyles"`
	Dxfs         *XDxfs             `xml:"dxfs"`
	TableStyles  *XTableStyles      `xml:"tableStyles"`
	ExtLst       *XStyleSheetExtLst `xml:"extLst"`
}

// XFonts Fonts type
type XFonts struct {
	Count      int    `xml:"count,attr"`
	KnownFonts string `xml:"x14ac:knownFonts,attr"`

	Font []*XStyleSheetFont `xml:"font"`
}

// XStyleSheetFont StyleSheetFont type
type XStyleSheetFont struct {
	Sz     *XValAttrElement `xml:"sz"`
	Color  *XColor          `xml:"color"`
	Name   *XValAttrElement `xml:"name"`
	Family *XValAttrElement `xml:"family"`
	Scheme *XValAttrElement `xml:"scheme"`
	B      *XValAttrElement `xml:"b,omitempty"`
	I      *XValAttrElement `xml:"i,omitempty"`
	U      *XValAttrElement `xml:"u,omitempty"`
	Strike *XValAttrElement `xml:"strike,omitempty"`
}

// XColor Color type
type XColor struct {
	Theme string `xml:"theme,attr"`
}

// XFills Fills type
type XFills struct {
	Count int `xml:"count,attr"`

	Fill []*XFill `xml:"fill"`
}

// XFill Fill type
type XFill struct {
	PatternFill *XPatternFill `xml:"patternFill"`
}

// XPatternFill PatternFill type
type XPatternFill struct {
	PatternType string `xml:"patternType,attr"`
}

// XBorders Borders type
type XBorders struct {
	Count int `xml:"count,attr"`

	Border []*XBorder `xml:"border"`
}

// XBorder Border type
type XBorder struct {
	Left     string `xml:"left"`
	Right    string `xml:"right"`
	Top      string `xml:"top"`
	Bottom   string `xml:"bottom"`
	Diagonal string `xml:"diagonal"`
}

// XCellStyleXfs CellStyleXfs type
type XCellStyleXfs struct {
	Count int `xml:"count,attr"`

	Xf []*XXf `xml:"xf"`
}

// XXf Xf type
type XXf struct {
	NumFmtID string `xml:"numFmtId,attr"`
	FontID   string `xml:"fontId,attr"`
	FillID   string `xml:"fillId,attr"`
	BorderID string `xml:"borderId,attr"`
	XfID     string `xml:"xfId,attr,omitempty"`
}

// XCellXfs CellXfs type
type XCellXfs struct {
	Count int `xml:"count,attr"`

	Xf []*XXf `xml:"xf"`
}

// XCellStyles CellStyles type
type XCellStyles struct {
	Count int `xml:"count,attr"`

	CellStyle []*XCellStyle `xml:"cellStyle"`
}

// XCellStyle CellStyle type
type XCellStyle struct {
	Name      string `xml:"name,attr"`
	XfID      string `xml:"xfId,attr"`
	BuiltinID string `xml:"builtinId,attr"`
}

// XDxfs Dxfs type
type XDxfs struct {
	Count int `xml:"count,attr"`
}

// XTableStyles TableStyles type
type XTableStyles struct {
	Count             int    `xml:"count,attr"`
	DefaultTableStyle string `xml:"defaultTableStyle,attr"`
	DefaultPivotStyle string `xml:"defaultPivotStyle,attr"`
}

// XStyleSheetExtLst StyleSheetExtLst type
type XStyleSheetExtLst struct {
	Ext []*XStyleSheetExt `xml:"ext"`
}

// XStyleSheetExt StyleSheetExt type
type XStyleSheetExt struct {
	URI            string           `xml:"uri,attr"`
	XmlnsX14       string           `xml:"xmlns:x14,attr,omitempty"`
	XmlnsX15       string           `xml:"xmlns:x15,attr,omitempty"`
	SlicerStyles   *XSlicerStyles   `xml:"x14:slicerStyles,omitempty"`
	TimelineStyles *XTimelineStyles `xml:"x15:timelineStyles,omitempty"`
}

// XSlicerStyles SlicerStyles type
type XSlicerStyles struct {
	DefaultSlicerStyle string `xml:"defaultSlicerStyle,attr"`
}

// XTimelineStyles TimelineStyles type
type XTimelineStyles struct {
	DefaultTimelineStyle string `xml:"defaultTimelineStyle,attr"`
}

// NewDefaultXStyleSheet create *XStyleSheet with default template
func NewDefaultXStyleSheet() *XStyleSheet {
	return &XStyleSheet{
		XmlnsMc:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		McIgnorable: "x14ac x16r2",
		XmlnsX14ac:  "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
		XmlnsX16r2:  "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",

		Fonts: &XFonts{
			Count:      1,
			KnownFonts: "1",
			Font: []*XStyleSheetFont{
				{
					Sz:     &XValAttrElement{Val: "11"},
					Color:  &XColor{Theme: "1"},
					Name:   &XValAttrElement{Val: "等线"},
					Family: &XValAttrElement{Val: "2"},
					Scheme: &XValAttrElement{Val: "minor"},
				},
			},
		},
		Fills: &XFills{
			Count: 2,
			Fill: []*XFill{
				{PatternFill: &XPatternFill{PatternType: "none"}},
				{PatternFill: &XPatternFill{PatternType: "gray125"}},
			},
		},
		Borders: &XBorders{
			Count: 1,
			Border: []*XBorder{
				{
					Left:     "",
					Right:    "",
					Top:      "",
					Bottom:   "",
					Diagonal: "",
				},
			},
		},
		CellStyleXfs: &XCellStyleXfs{
			Count: 1,
			Xf: []*XXf{
				{
					NumFmtID: "0",
					FontID:   "0",
					FillID:   "0",
					BorderID: "0",
				},
			},
		},
		CellXfs: &XCellXfs{
			Count: 1,
			Xf: []*XXf{
				{
					NumFmtID: "0",
					FontID:   "0",
					FillID:   "0",
					BorderID: "0",
					XfID:     "0",
				},
			},
		},
		CellStyles: &XCellStyles{
			Count: 1,
			CellStyle: []*XCellStyle{
				{
					Name:      "Normal",
					XfID:      "0",
					BuiltinID: "0",
				},
			},
		},
		Dxfs: &XDxfs{Count: 0},
		TableStyles: &XTableStyles{
			Count:             0,
			DefaultTableStyle: "TableStyleMedium2",
			DefaultPivotStyle: "PivotStyleLight16",
		},
		ExtLst: &XStyleSheetExtLst{
			Ext: []*XStyleSheetExt{
				{
					URI:          "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
					XmlnsX14:     "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
					SlicerStyles: &XSlicerStyles{DefaultSlicerStyle: "SlicerStyleLight1"},
				},
				{
					URI:            "{9260A510-F301-46a8-8635-F512D64BE5F5}",
					XmlnsX15:       "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
					TimelineStyles: &XTimelineStyles{DefaultTimelineStyle: "TimeSlicerStyleLight1"},
				},
			},
		},
	}
}
