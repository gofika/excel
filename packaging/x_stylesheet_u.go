package packaging

import "encoding/xml"

// XStyleSheetU fix XML ns for XStyleSheet
type XStyleSheetU struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	XmlnsMc     string   `xml:"mc,attr"`
	McIgnorable string   `xml:"Ignorable,attr"`
	XmlnsX14ac  string   `xml:"x14ac,attr"`
	XmlnsX16r2  string   `xml:"x16r2,attr"`

	Fonts        *XFontsU            `xml:"fonts"`
	Fills        *XFillsU            `xml:"fills"`
	Borders      *XBordersU          `xml:"borders"`
	CellStyleXfs *XCellStyleXfsU     `xml:"cellStyleXfs"`
	CellXfs      *XCellXfsU          `xml:"cellXfs"`
	CellStyles   *XCellStylesU       `xml:"cellStyles"`
	Dxfs         *XDxfsU             `xml:"dxfs"`
	TableStyles  *XTableStylesU      `xml:"tableStyles"`
	ExtLst       *XStyleSheetExtLstU `xml:"extLst"`
}

// XFontsU fix XML ns for XFonts
type XFontsU struct {
	Count      int    `xml:"count,attr"`
	KnownFonts string `xml:"knownFonts,attr"`

	Font []*XStyleSheetFontU `xml:"font"`
}

// XStyleSheetFontU fix XML ns for XStyleSheetFont
type XStyleSheetFontU struct {
	Sz     *XValAttrElementU `xml:"sz"`
	Color  *XColorU          `xml:"color"`
	Name   *XValAttrElementU `xml:"name"`
	Family *XValAttrElementU `xml:"family"`
	Scheme *XValAttrElementU `xml:"scheme"`
	B      *XValAttrElementU `xml:"b,omitempty"`
	I      *XValAttrElementU `xml:"i,omitempty"`
	U      *XValAttrElementU `xml:"u,omitempty"`
	Strike *XValAttrElementU `xml:"strike,omitempty"`
}

// XColorU fix XML ns for XColor
type XColorU struct {
	Theme string `xml:"theme,attr"`
}

// XFillsU fix XML ns for XFills
type XFillsU struct {
	Count int `xml:"count,attr"`

	Fill []*XFillU `xml:"fill"`
}

// XFillU fix XML ns for XFill
type XFillU struct {
	PatternFill *XPatternFillU `xml:"patternFill"`
}

// XPatternFillU fix XML ns for XPatternFill
type XPatternFillU struct {
	PatternType string `xml:"patternType,attr"`
}

// XBordersU fix XML ns for XBorders
type XBordersU struct {
	Count int `xml:"count,attr"`

	Border []*XBorderU `xml:"border"`
}

// XBorderU fix XML ns for XBorder
type XBorderU struct {
	Left     string `xml:"left"`
	Right    string `xml:"right"`
	Top      string `xml:"top"`
	Bottom   string `xml:"bottom"`
	Diagonal string `xml:"diagonal"`
}

// XCellStyleXfsU fix XML ns for XCellStyleXfs
type XCellStyleXfsU struct {
	Count int `xml:"count,attr"`

	Xf []*XXfU `xml:"xf"`
}

// XXfU fix XML ns for XXf
type XXfU struct {
	NumFmtID string `xml:"numFmtId,attr"`
	FontID   string `xml:"fontId,attr"`
	FillID   string `xml:"fillId,attr"`
	BorderID string `xml:"borderId,attr"`
	XfID     string `xml:"xfId,attr,omitempty"`
}

// XCellXfsU fix XML ns for XCellXfs
type XCellXfsU struct {
	Count int `xml:"count,attr"`

	Xf []*XXfU `xml:"xf"`
}

// XCellStylesU fix XML ns for XCellStyles
type XCellStylesU struct {
	Count int `xml:"count,attr"`

	CellStyle []*XCellStyleU `xml:"cellStyle"`
}

// XCellStyleU fix XML ns for XCellStyle
type XCellStyleU struct {
	Name      string `xml:"name,attr"`
	XfID      string `xml:"xfId,attr"`
	BuiltinID string `xml:"builtinId,attr"`
}

// XDxfsU fix XML ns for XDxfs
type XDxfsU struct {
	Count int `xml:"count,attr"`
}

// XTableStylesU fix XML ns for XTableStyles
type XTableStylesU struct {
	Count             int    `xml:"count,attr"`
	DefaultTableStyle string `xml:"defaultTableStyle,attr"`
	DefaultPivotStyle string `xml:"defaultPivotStyle,attr"`
}

// XStyleSheetExtLstU fix XML ns for XStyleSheetExtLst
type XStyleSheetExtLstU struct {
	Ext []*XStyleSheetExtU `xml:"ext"`
}

// XStyleSheetExtU fix XML ns for XStyleSheetExt
type XStyleSheetExtU struct {
	URI            string            `xml:"uri,attr"`
	XmlnsX14       string            `xml:"x14,attr,omitempty"`
	XmlnsX15       string            `xml:"x15,attr,omitempty"`
	SlicerStyles   *XSlicerStylesU   `xml:"slicerStyles,omitempty"`
	TimelineStyles *XTimelineStylesU `xml:"timelineStyles,omitempty"`
}

// XSlicerStylesU fix XML ns for XSlicerStyles
type XSlicerStylesU struct {
	DefaultSlicerStyle string `xml:"defaultSlicerStyle,attr"`
}

// XTimelineStylesU fix XML ns for XTimelineStyles
type XTimelineStylesU struct {
	DefaultTimelineStyle string `xml:"defaultTimelineStyle,attr"`
}
