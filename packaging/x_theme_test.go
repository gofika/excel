package packaging

import (
	"fmt"
	"github.com/gofika/util/fileutil"
	"path"

	. "gopkg.in/check.v1"
)

const defaultThemeTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1>
                <a:sysClr val="windowText" lastClr="000000"></a:sysClr>
            </a:dk1>
            <a:lt1>
                <a:sysClr val="window" lastClr="FFFFFF"></a:sysClr>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="44546A"></a:srgbClr>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="E7E6E6"></a:srgbClr>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="5B9BD5"></a:srgbClr>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="ED7D31"></a:srgbClr>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="A5A5A5"></a:srgbClr>
            </a:accent3>
            <a:accent4>
                <a:srgbClr val="FFC000"></a:srgbClr>
            </a:accent4>
            <a:accent5>
                <a:srgbClr val="4472C4"></a:srgbClr>
            </a:accent5>
            <a:accent6>
                <a:srgbClr val="70AD47"></a:srgbClr>
            </a:accent6>
            <a:hlink>
                <a:srgbClr val="0563C1"></a:srgbClr>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="954F72"></a:srgbClr>
            </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Calibri Light" panose="020F0302020204030204"></a:latin>
                <a:ea typeface=""></a:ea>
                <a:cs typeface=""></a:cs>
                <a:font script="Jpan" typeface="Yu Gothic Light"></a:font>
                <a:font script="Hang" typeface="맑은 고딕"></a:font>
                <a:font script="Hans" typeface="等线 Light"></a:font>
                <a:font script="Hant" typeface="新細明體"></a:font>
                <a:font script="Arab" typeface="Times New Roman"></a:font>
                <a:font script="Hebr" typeface="Times New Roman"></a:font>
                <a:font script="Thai" typeface="Tahoma"></a:font>
                <a:font script="Ethi" typeface="Nyala"></a:font>
                <a:font script="Beng" typeface="Vrinda"></a:font>
                <a:font script="Gujr" typeface="Shruti"></a:font>
                <a:font script="Khmr" typeface="MoolBoran"></a:font>
                <a:font script="Knda" typeface="Tunga"></a:font>
                <a:font script="Guru" typeface="Raavi"></a:font>
                <a:font script="Cans" typeface="Euphemia"></a:font>
                <a:font script="Cher" typeface="Plantagenet Cherokee"></a:font>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"></a:font>
                <a:font script="Tibt" typeface="Microsoft Himalaya"></a:font>
                <a:font script="Thaa" typeface="MV Boli"></a:font>
                <a:font script="Deva" typeface="Mangal"></a:font>
                <a:font script="Telu" typeface="Gautami"></a:font>
                <a:font script="Taml" typeface="Latha"></a:font>
                <a:font script="Syrc" typeface="Estrangelo Edessa"></a:font>
                <a:font script="Orya" typeface="Kalinga"></a:font>
                <a:font script="Mlym" typeface="Kartika"></a:font>
                <a:font script="Laoo" typeface="DokChampa"></a:font>
                <a:font script="Sinh" typeface="Iskoola Pota"></a:font>
                <a:font script="Mong" typeface="Mongolian Baiti"></a:font>
                <a:font script="Viet" typeface="Times New Roman"></a:font>
                <a:font script="Uigh" typeface="Microsoft Uighur"></a:font>
                <a:font script="Geor" typeface="Sylfaen"></a:font>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri" panose="020F0502020204030204"></a:latin>
                <a:ea typeface=""></a:ea>
                <a:cs typeface=""></a:cs>
                <a:font script="Jpan" typeface="Yu Gothic"></a:font>
                <a:font script="Hang" typeface="맑은 고딕"></a:font>
                <a:font script="Hans" typeface="等线"></a:font>
                <a:font script="Hant" typeface="新細明體"></a:font>
                <a:font script="Arab" typeface="Arial"></a:font>
                <a:font script="Hebr" typeface="Arial"></a:font>
                <a:font script="Thai" typeface="Tahoma"></a:font>
                <a:font script="Ethi" typeface="Nyala"></a:font>
                <a:font script="Beng" typeface="Vrinda"></a:font>
                <a:font script="Gujr" typeface="Shruti"></a:font>
                <a:font script="Khmr" typeface="DaunPenh"></a:font>
                <a:font script="Knda" typeface="Tunga"></a:font>
                <a:font script="Guru" typeface="Raavi"></a:font>
                <a:font script="Cans" typeface="Euphemia"></a:font>
                <a:font script="Cher" typeface="Plantagenet Cherokee"></a:font>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"></a:font>
                <a:font script="Tibt" typeface="Microsoft Himalaya"></a:font>
                <a:font script="Thaa" typeface="MV Boli"></a:font>
                <a:font script="Deva" typeface="Mangal"></a:font>
                <a:font script="Telu" typeface="Gautami"></a:font>
                <a:font script="Taml" typeface="Latha"></a:font>
                <a:font script="Syrc" typeface="Estrangelo Edessa"></a:font>
                <a:font script="Orya" typeface="Kalinga"></a:font>
                <a:font script="Mlym" typeface="Kartika"></a:font>
                <a:font script="Laoo" typeface="DokChampa"></a:font>
                <a:font script="Sinh" typeface="Iskoola Pota"></a:font>
                <a:font script="Mong" typeface="Mongolian Baiti"></a:font>
                <a:font script="Viet" typeface="Arial"></a:font>
                <a:font script="Uigh" typeface="Microsoft Uighur"></a:font>
                <a:font script="Geor" typeface="Sylfaen"></a:font>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"></a:schemeClr>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="110000"></a:lumMod>
                                <a:satMod val="105000"></a:satMod>
                                <a:tint val="67000"></a:tint>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="105000"></a:lumMod>
                                <a:satMod val="103000"></a:satMod>
                                <a:tint val="73000"></a:tint>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="105000"></a:lumMod>
                                <a:satMod val="109000"></a:satMod>
                                <a:tint val="81000"></a:tint>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"></a:lin>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="102000"></a:lumMod>
                                <a:satMod val="103000"></a:satMod>
                                <a:tint val="94000"></a:tint>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="100000"></a:lumMod>
                                <a:satMod val="110000"></a:satMod>
                                <a:shade val="100000"></a:shade>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="99000"></a:lumMod>
                                <a:satMod val="120000"></a:satMod>
                                <a:shade val="78000"></a:shade>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"></a:lin>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"></a:schemeClr>
                    </a:solidFill>
                    <a:prstDash val="solid"></a:prstDash>
                    <a:miter lim="800000"></a:miter>
                </a:ln>
                <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"></a:schemeClr>
                    </a:solidFill>
                    <a:prstDash val="solid"></a:prstDash>
                    <a:miter lim="800000"></a:miter>
                </a:ln>
                <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"></a:schemeClr>
                    </a:solidFill>
                    <a:prstDash val="solid"></a:prstDash>
                    <a:miter lim="800000"></a:miter>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst></a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst></a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="63000"></a:alpha>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"></a:schemeClr>
                </a:solidFill>
                <a:solidFill>
                    <a:schemeClr val="phClr">
                        <a:satMod val="170000"></a:satMod>
                        <a:tint val="95000"></a:tint>
                    </a:schemeClr>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="102000"></a:lumMod>
                                <a:satMod val="150000"></a:satMod>
                                <a:tint val="93000"></a:tint>
                                <a:shade val="98000"></a:shade>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="103000"></a:lumMod>
                                <a:satMod val="130000"></a:satMod>
                                <a:tint val="98000"></a:tint>
                                <a:shade val="90000"></a:shade>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:satMod val="120000"></a:satMod>
                                <a:shade val="63000"></a:shade>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"></a:lin>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults></a:objectDefaults>
    <a:extraClrSchemeLst></a:extraClrSchemeLst>
    <a:extLst>
        <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
            <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"></thm15:themeFamily>
        </a:ext>
    </a:extLst>
</a:theme>`

func testTheme(c *C, file *XFile) {
	var result string
	var err error
	result, err = XMLMarshalAppendHeadIndent(file.Themes[0])
	c.Assert(err, IsNil)
	if needWriteTestFile {
		err = fileutil.WriteFile(path.Join(templatePath, ThemePath, fmt.Sprintf(ThemeFileName, 1)), []byte(result))
		c.Assert(err, IsNil)
	}
	c.Assert(result, Equals, defaultThemeTemplate)
}
