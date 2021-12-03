package xlsx

import "encoding/xml"

// xlsxChartsheet directly maps the chartsheet element of Chartsheet Parts in
// a SpreadsheetML document.
type xlsxChartsheet struct {
	XMLName          xml.Name                   `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main chartsheet"`
	SheetPr          *xlsxChartsheetPr          `xml:"sheetPr"`
	SheetViews       *xlsxChartsheetViews       `xml:"sheetViews"`
	SheetProtection  *xlsxChartsheetProtection  `xml:"sheetProtection"`
	CustomSheetViews *xlsxCustomChartsheetViews `xml:"customSheetViews"`
	PageMargins      *xlsxPageMargins           `xml:"pageMargins"`
	PageSetup        *xlsxPageSetUp             `xml:"pageSetup"`
	HeaderFooter     *xlsxHeaderFooter          `xml:"headerFooter"`
	Drawing          *xlsxDrawing               `xml:"drawing"`
	DrawingHF        *xlsxDrawingHF             `xml:"drawingHF"`
	Picture          *xlsxPicture               `xml:"picture"`
	WebPublishItems  *xlsxInnerXML              `xml:"webPublishItems"`
	ExtLst           *xlsxExtLst                `xml:"extLst"`
}

// xlsxChartsheetPr specifies chart sheet properties.
type xlsxChartsheetPr struct {
	XMLName       xml.Name      `xml:"sheetPr"`
	PublishedAttr bool          `xml:"published,attr,omitempty"`
	CodeNameAttr  string        `xml:"codeName,attr,omitempty"`
	TabColor      *xlsxTabColor `xml:"tabColor"`
}

// xlsxChartsheetViews specifies chart sheet views.
type xlsxChartsheetViews struct {
	XMLName   xml.Name              `xml:"sheetViews"`
	SheetView []*xlsxChartsheetView `xml:"sheetView"`
	ExtLst    []*xlsxExtLst         `xml:"extLst"`
}

// xlsxChartsheetView defines custom view properties for chart sheets.
type xlsxChartsheetView struct {
	XMLName            xml.Name      `xml:"sheetView"`
	TabSelectedAttr    bool          `xml:"tabSelected,attr,omitempty"`
	ZoomScaleAttr      uint32        `xml:"zoomScale,attr,omitempty"`
	WorkbookViewIDAttr uint32        `xml:"workbookViewId,attr"`
	ZoomToFitAttr      bool          `xml:"zoomToFit,attr,omitempty"`
	ExtLst             []*xlsxExtLst `xml:"extLst"`
}

// xlsxChartsheetProtection collection expresses the chart sheet protection
// options to enforce when the chart sheet is protected.
type xlsxChartsheetProtection struct {
	XMLName           xml.Name `xml:"sheetProtection"`
	AlgorithmNameAttr string   `xml:"algorithmName,attr,omitempty"`
	HashValueAttr     []byte   `xml:"hashValue,attr,omitempty"`
	SaltValueAttr     []byte   `xml:"saltValue,attr,omitempty"`
	SpinCountAttr     uint32   `xml:"spinCount,attr,omitempty"`
	ContentAttr       bool     `xml:"content,attr,omitempty"`
	ObjectsAttr       bool     `xml:"objects,attr,omitempty"`
}

// xlsxCustomChartsheetViews collection of custom Chart Sheet View
// information.
type xlsxCustomChartsheetViews struct {
	XMLName         xml.Name                    `xml:"customSheetViews"`
	CustomSheetView []*xlsxCustomChartsheetView `xml:"customSheetView"`
}

// xlsxCustomChartsheetView defines custom view properties for chart sheets.
type xlsxCustomChartsheetView struct {
	XMLName       xml.Name            `xml:"customSheetView"`
	GUIDAttr      string              `xml:"guid,attr"`
	ScaleAttr     uint32              `xml:"scale,attr,omitempty"`
	StateAttr     string              `xml:"state,attr,omitempty"`
	ZoomToFitAttr bool                `xml:"zoomToFit,attr,omitempty"`
	PageMargins   []*xlsxPageMargins  `xml:"pageMargins"`
	PageSetup     []*xlsxPageSetUp    `xml:"pageSetup"`
	HeaderFooter  []*xlsxHeaderFooter `xml:"headerFooter"`
}
