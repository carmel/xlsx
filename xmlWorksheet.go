package xlsx

import (
	"encoding/xml"
	"sync"
)

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main.
type xlsxWorksheet struct {
	sync.Mutex
	XMLName               xml.Name                     `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetPr               *xlsxSheetPr                 `xml:"sheetPr"`
	Dimension             *xlsxDimension               `xml:"dimension"`
	SheetViews            *xlsxSheetViews              `xml:"sheetViews"`
	SheetFormatPr         *xlsxSheetFormatPr           `xml:"sheetFormatPr"`
	Cols                  *xlsxCols                    `xml:"cols"`
	SheetData             xlsxSheetData                `xml:"sheetData"`
	SheetCalcPr           *xlsxInnerXML                `xml:"sheetCalcPr"`
	SheetProtection       *xlsxSheetProtection         `xml:"sheetProtection"`
	ProtectedRanges       *xlsxInnerXML                `xml:"protectedRanges"`
	Scenarios             *xlsxInnerXML                `xml:"scenarios"`
	AutoFilter            *xlsxAutoFilter              `xml:"autoFilter"`
	SortState             *xlsxSortState               `xml:"sortState"`
	DataConsolidate       *xlsxInnerXML                `xml:"dataConsolidate"`
	CustomSheetViews      *xlsxCustomSheetViews        `xml:"customSheetViews"`
	MergeCells            *xlsxMergeCells              `xml:"mergeCells"`
	PhoneticPr            *xlsxPhoneticPr              `xml:"phoneticPr"`
	ConditionalFormatting []*xlsxConditionalFormatting `xml:"conditionalFormatting"`
	DataValidations       *xlsxDataValidations         `xml:"dataValidations"`
	Hyperlinks            *xlsxHyperlinks              `xml:"hyperlinks"`
	PrintOptions          *xlsxPrintOptions            `xml:"printOptions"`
	PageMargins           *xlsxPageMargins             `xml:"pageMargins"`
	PageSetUp             *xlsxPageSetUp               `xml:"pageSetup"`
	HeaderFooter          *xlsxHeaderFooter            `xml:"headerFooter"`
	RowBreaks             *xlsxBreaks                  `xml:"rowBreaks"`
	ColBreaks             *xlsxBreaks                  `xml:"colBreaks"`
	CustomProperties      *xlsxInnerXML                `xml:"customProperties"`
	CellWatches           *xlsxInnerXML                `xml:"cellWatches"`
	IgnoredErrors         *xlsxInnerXML                `xml:"ignoredErrors"`
	SmartTags             *xlsxInnerXML                `xml:"smartTags"`
	Drawing               *xlsxDrawing                 `xml:"drawing"`
	LegacyDrawing         *xlsxLegacyDrawing           `xml:"legacyDrawing"`
	LegacyDrawingHF       *xlsxLegacyDrawingHF         `xml:"legacyDrawingHF"`
	DrawingHF             *xlsxDrawingHF               `xml:"drawingHF"`
	Picture               *xlsxPicture                 `xml:"picture"`
	OleObjects            *xlsxInnerXML                `xml:"oleObjects"`
	Controls              *xlsxInnerXML                `xml:"controls"`
	WebPublishItems       *xlsxInnerXML                `xml:"webPublishItems"`
	TableParts            *xlsxTableParts              `xml:"tableParts"`
	ExtLst                *xlsxExtLst                  `xml:"extLst"`
}

// xlsxDrawing change r:id to rid in the namespace.
type xlsxDrawing struct {
	XMLName xml.Name `xml:"drawing"`
	RID     string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// xlsxHeaderFooter directly maps the headerFooter element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - When printed or
// viewed in page layout view (§18.18.69), each page of a worksheet can have a
// page header, a page footer, or both. The headers and footers on odd-numbered
// pages can differ from those on even-numbered pages, and the headers and
// footers on the first page can differ from those on odd- and even-numbered
// pages. In the latter case, the first page is not considered an odd page.
type xlsxHeaderFooter struct {
	XMLName          xml.Name       `xml:"headerFooter"`
	AlignWithMargins bool           `xml:"alignWithMargins,attr,omitempty"`
	DifferentFirst   bool           `xml:"differentFirst,attr,omitempty"`
	DifferentOddEven bool           `xml:"differentOddEven,attr,omitempty"`
	ScaleWithDoc     bool           `xml:"scaleWithDoc,attr,omitempty"`
	OddHeader        string         `xml:"oddHeader,omitempty"`
	OddFooter        string         `xml:"oddFooter,omitempty"`
	EvenHeader       string         `xml:"evenHeader,omitempty"`
	EvenFooter       string         `xml:"evenFooter,omitempty"`
	FirstFooter      string         `xml:"firstFooter,omitempty"`
	FirstHeader      string         `xml:"firstHeader,omitempty"`
	DrawingHF        *xlsxDrawingHF `xml:"drawingHF"`
}

// xlsxDrawingHF (Drawing Reference in Header Footer) specifies the usage of
// drawing objects to be rendered in the headers and footers of the sheet. It
// specifies an explicit relationship to the part containing the DrawingML
// shapes used in the headers and footers. It also indicates where in the
// headers and footers each shape belongs. One drawing object can appear in
// each of the left section, center section and right section of a header and
// a footer.
type xlsxDrawingHF struct {
	Content string `xml:",innerxml"`
}

// xlsxPageSetUp directly maps the pageSetup element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - Page setup
// settings for the worksheet.
type xlsxPageSetUp struct {
	XMLName            xml.Name `xml:"pageSetup"`
	BlackAndWhite      bool     `xml:"blackAndWhite,attr,omitempty"`
	CellComments       string   `xml:"cellComments,attr,omitempty"`
	Copies             int      `xml:"copies,attr,omitempty"`
	Draft              bool     `xml:"draft,attr,omitempty"`
	Errors             string   `xml:"errors,attr,omitempty"`
	FirstPageNumber    string   `xml:"firstPageNumber,attr,omitempty"`
	FitToHeight        int      `xml:"fitToHeight,attr,omitempty"`
	FitToWidth         int      `xml:"fitToWidth,attr,omitempty"`
	HorizontalDPI      int      `xml:"horizontalDpi,attr,omitempty"`
	RID                string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	Orientation        string   `xml:"orientation,attr,omitempty"`
	PageOrder          string   `xml:"pageOrder,attr,omitempty"`
	PaperHeight        string   `xml:"paperHeight,attr,omitempty"`
	PaperSize          int      `xml:"paperSize,attr,omitempty"`
	PaperWidth         string   `xml:"paperWidth,attr,omitempty"`
	Scale              int      `xml:"scale,attr,omitempty"`
	UseFirstPageNumber bool     `xml:"useFirstPageNumber,attr,omitempty"`
	UsePrinterDefaults bool     `xml:"usePrinterDefaults,attr,omitempty"`
	VerticalDPI        int      `xml:"verticalDpi,attr,omitempty"`
}

// xlsxPrintOptions directly maps the printOptions element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - Print options for
// the sheet. Printer-specific settings are stored separately in the Printer
// Settings part.
type xlsxPrintOptions struct {
	XMLName            xml.Name `xml:"printOptions"`
	GridLines          bool     `xml:"gridLines,attr,omitempty"`
	GridLinesSet       bool     `xml:"gridLinesSet,attr,omitempty"`
	Headings           bool     `xml:"headings,attr,omitempty"`
	HorizontalCentered bool     `xml:"horizontalCentered,attr,omitempty"`
	VerticalCentered   bool     `xml:"verticalCentered,attr,omitempty"`
}

// xlsxPageMargins directly maps the pageMargins element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - Page margins for
// a sheet or a custom sheet view.
type xlsxPageMargins struct {
	XMLName xml.Name `xml:"pageMargins"`
	Bottom  float64  `xml:"bottom,attr"`
	Footer  float64  `xml:"footer,attr"`
	Header  float64  `xml:"header,attr"`
	Left    float64  `xml:"left,attr"`
	Right   float64  `xml:"right,attr"`
	Top     float64  `xml:"top,attr"`
}

// xlsxSheetFormatPr directly maps the sheetFormatPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main. This element
// specifies the sheet formatting properties.
type xlsxSheetFormatPr struct {
	XMLName          xml.Name `xml:"sheetFormatPr"`
	BaseColWidth     uint8    `xml:"baseColWidth,attr,omitempty"`
	DefaultColWidth  float64  `xml:"defaultColWidth,attr,omitempty"`
	DefaultRowHeight float64  `xml:"defaultRowHeight,attr"`
	CustomHeight     bool     `xml:"customHeight,attr,omitempty"`
	ZeroHeight       bool     `xml:"zeroHeight,attr,omitempty"`
	ThickTop         bool     `xml:"thickTop,attr,omitempty"`
	ThickBottom      bool     `xml:"thickBottom,attr,omitempty"`
	OutlineLevelRow  uint8    `xml:"outlineLevelRow,attr,omitempty"`
	OutlineLevelCol  uint8    `xml:"outlineLevelCol,attr,omitempty"`
}

// xlsxSheetViews represents worksheet views collection.
type xlsxSheetViews struct {
	XMLName   xml.Name        `xml:"sheetViews"`
	SheetView []xlsxSheetView `xml:"sheetView"`
}

// xlsxSheetView represents a single sheet view definition. When more than one
// sheet view is defined in the file, it means that when opening the workbook,
// each sheet view corresponds to a separate window within the spreadsheet
// application, where each window is showing the particular sheet containing
// the same workbookViewId value, the last sheetView definition is loaded, and
// the others are discarded. When multiple windows are viewing the same sheet,
// multiple sheetView elements (with corresponding workbookView entries) are
// saved.
type xlsxSheetView struct {
	WindowProtection         bool             `xml:"windowProtection,attr,omitempty"`
	ShowFormulas             bool             `xml:"showFormulas,attr,omitempty"`
	ShowGridLines            *bool            `xml:"showGridLines,attr"`
	ShowRowColHeaders        *bool            `xml:"showRowColHeaders,attr"`
	ShowZeros                *bool            `xml:"showZeros,attr,omitempty"`
	RightToLeft              bool             `xml:"rightToLeft,attr,omitempty"`
	TabSelected              bool             `xml:"tabSelected,attr,omitempty"`
	ShowWhiteSpace           *bool            `xml:"showWhiteSpace,attr"`
	ShowOutlineSymbols       bool             `xml:"showOutlineSymbols,attr,omitempty"`
	DefaultGridColor         *bool            `xml:"defaultGridColor,attr"`
	View                     string           `xml:"view,attr,omitempty"`
	TopLeftCell              string           `xml:"topLeftCell,attr,omitempty"`
	ColorID                  int              `xml:"colorId,attr,omitempty"`
	ZoomScale                float64          `xml:"zoomScale,attr,omitempty"`
	ZoomScaleNormal          float64          `xml:"zoomScaleNormal,attr,omitempty"`
	ZoomScalePageLayoutView  float64          `xml:"zoomScalePageLayoutView,attr,omitempty"`
	ZoomScaleSheetLayoutView float64          `xml:"zoomScaleSheetLayoutView,attr,omitempty"`
	WorkbookViewID           int              `xml:"workbookViewId,attr"`
	Pane                     *xlsxPane        `xml:"pane,omitempty"`
	Selection                []*xlsxSelection `xml:"selection"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - Worksheet view
// selection.
type xlsxSelection struct {
	ActiveCell   string `xml:"activeCell,attr,omitempty"`
	ActiveCellID *int   `xml:"activeCellId,attr"`
	Pane         string `xml:"pane,attr,omitempty"`
	SQRef        string `xml:"sqref,attr,omitempty"`
}

// xlsxSelection directly maps the selection element. Worksheet view pane.
type xlsxPane struct {
	ActivePane  string  `xml:"activePane,attr,omitempty"`
	State       string  `xml:"state,attr,omitempty"` // Either "split" or "frozen"
	TopLeftCell string  `xml:"topLeftCell,attr,omitempty"`
	XSplit      float64 `xml:"xSplit,attr,omitempty"`
	YSplit      float64 `xml:"ySplit,attr,omitempty"`
}

// xlsxSheetPr directly maps the sheetPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - Sheet-level
// properties.
type xlsxSheetPr struct {
	XMLName                           xml.Name         `xml:"sheetPr"`
	SyncHorizontal                    bool             `xml:"syncHorizontal,attr,omitempty"`
	SyncVertical                      bool             `xml:"syncVertical,attr,omitempty"`
	SyncRef                           string           `xml:"syncRef,attr,omitempty"`
	TransitionEvaluation              bool             `xml:"transitionEvaluation,attr,omitempty"`
	TransitionEntry                   bool             `xml:"transitionEntry,attr,omitempty"`
	Published                         *bool            `xml:"published,attr"`
	CodeName                          string           `xml:"codeName,attr,omitempty"`
	FilterMode                        bool             `xml:"filterMode,attr,omitempty"`
	EnableFormatConditionsCalculation *bool            `xml:"enableFormatConditionsCalculation,attr"`
	TabColor                          *xlsxTabColor    `xml:"tabColor,omitempty"`
	OutlinePr                         *xlsxOutlinePr   `xml:"outlinePr,omitempty"`
	PageSetUpPr                       *xlsxPageSetUpPr `xml:"pageSetUpPr,omitempty"`
}

// xlsxOutlinePr maps to the outlinePr element. SummaryBelow allows you to
// adjust the direction of grouper controls.
type xlsxOutlinePr struct {
	ApplyStyles        *bool `xml:"applyStyles,attr"`
	SummaryBelow       bool  `xml:"summaryBelow,attr"`
	SummaryRight       bool  `xml:"summaryRight,attr"`
	ShowOutlineSymbols bool  `xml:"showOutlineSymbols,attr"`
}

// xlsxPageSetUpPr expresses page setup properties of the worksheet.
type xlsxPageSetUpPr struct {
	AutoPageBreaks bool `xml:"autoPageBreaks,attr,omitempty"`
	FitToPage      bool `xml:"fitToPage,attr,omitempty"`
}

// xlsxTabColor represents background color of the sheet tab.
type xlsxTabColor struct {
	Auto    bool    `xml:"auto,attr,omitempty"`
	Indexed int     `xml:"indexed,attr,omitempty"`
	RGB     string  `xml:"rgb,attr,omitempty"`
	Theme   int     `xml:"theme,attr,omitempty"`
	Tint    float64 `xml:"tint,attr,omitempty"`
}

// xlsxCols defines column width and column formatting for one or more columns
// of the worksheet.
type xlsxCols struct {
	XMLName xml.Name  `xml:"cols"`
	Col     []xlsxCol `xml:"col"`
}

// xlsxCol directly maps the col (Column Width & Formatting). Defines column
// width and column formatting for one or more columns of the worksheet.
type xlsxCol struct {
	BestFit      bool    `xml:"bestFit,attr,omitempty"`
	Collapsed    bool    `xml:"collapsed,attr,omitempty"`
	CustomWidth  bool    `xml:"customWidth,attr,omitempty"`
	Hidden       bool    `xml:"hidden,attr,omitempty"`
	Max          int     `xml:"max,attr"`
	Min          int     `xml:"min,attr"`
	OutlineLevel uint8   `xml:"outlineLevel,attr,omitempty"`
	Phonetic     bool    `xml:"phonetic,attr,omitempty"`
	Style        int     `xml:"style,attr,omitempty"`
	Width        float64 `xml:"width,attr,omitempty"`
}

// xlsxDimension directly maps the dimension element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - This element
// specifies the used range of the worksheet. It specifies the row and column
// bounds of used cells in the worksheet. This is optional and is not
// required. Used cells include cells with formulas, text content, and cell
// formatting. When an entire column is formatted, only the first cell in that
// column is considered used.
type xlsxDimension struct {
	XMLName xml.Name `xml:"dimension"`
	Ref     string   `xml:"ref,attr"`
}

// xlsxSheetData collection represents the cell table itself. This collection
// expresses information about each cell, grouped together by rows in the
// worksheet.
type xlsxSheetData struct {
	XMLName xml.Name  `xml:"sheetData"`
	Row     []xlsxRow `xml:"row"`
}

// xlsxRow directly maps the row element. The element expresses information
// about an entire row of a worksheet, and contains all cell definitions for a
// particular row in the worksheet.
type xlsxRow struct {
	C            []xlsxC `xml:"c"`
	R            int     `xml:"r,attr,omitempty"`
	Spans        string  `xml:"spans,attr,omitempty"`
	S            int     `xml:"s,attr,omitempty"`
	CustomFormat bool    `xml:"customFormat,attr,omitempty"`
	Ht           float64 `xml:"ht,attr,omitempty"`
	Hidden       bool    `xml:"hidden,attr,omitempty"`
	CustomHeight bool    `xml:"customHeight,attr,omitempty"`
	OutlineLevel uint8   `xml:"outlineLevel,attr,omitempty"`
	Collapsed    bool    `xml:"collapsed,attr,omitempty"`
	ThickTop     bool    `xml:"thickTop,attr,omitempty"`
	ThickBot     bool    `xml:"thickBot,attr,omitempty"`
	Ph           bool    `xml:"ph,attr,omitempty"`
}

// xlsxSortState directly maps the sortState element. This collection
// preserves the AutoFilter sort state.
type xlsxSortState struct {
	ColumnSort    bool   `xml:"columnSort,attr,omitempty"`
	CaseSensitive bool   `xml:"caseSensitive,attr,omitempty"`
	SortMethod    string `xml:"sortMethod,attr,omitempty"`
	Ref           string `xml:"ref,attr"`
	Content       string `xml:",innerxml"`
}

// xlsxCustomSheetViews directly maps the customSheetViews element. This is a
// collection of custom sheet views.
type xlsxCustomSheetViews struct {
	XMLName         xml.Name               `xml:"customSheetViews"`
	CustomSheetView []*xlsxCustomSheetView `xml:"customSheetView"`
}

// xlsxBrk directly maps the row or column break to use when paginating a
// worksheet.
type xlsxBrk struct {
	ID  int  `xml:"id,attr,omitempty"`
	Min int  `xml:"min,attr,omitempty"`
	Max int  `xml:"max,attr,omitempty"`
	Man bool `xml:"man,attr,omitempty"`
	Pt  bool `xml:"pt,attr,omitempty"`
}

// xlsxBreaks directly maps a collection of the row or column breaks.
type xlsxBreaks struct {
	Brk              []*xlsxBrk `xml:"brk"`
	Count            int        `xml:"count,attr,omitempty"`
	ManualBreakCount int        `xml:"manualBreakCount,attr,omitempty"`
}

// xlsxCustomSheetView directly maps the customSheetView element.
type xlsxCustomSheetView struct {
	Pane           *xlsxPane         `xml:"pane"`
	Selection      *xlsxSelection    `xml:"selection"`
	RowBreaks      *xlsxBreaks       `xml:"rowBreaks"`
	ColBreaks      *xlsxBreaks       `xml:"colBreaks"`
	PageMargins    *xlsxPageMargins  `xml:"pageMargins"`
	PrintOptions   *xlsxPrintOptions `xml:"printOptions"`
	PageSetup      *xlsxPageSetUp    `xml:"pageSetup"`
	HeaderFooter   *xlsxHeaderFooter `xml:"headerFooter"`
	AutoFilter     *xlsxAutoFilter   `xml:"autoFilter"`
	ExtLst         *xlsxExtLst       `xml:"extLst"`
	GUID           string            `xml:"guid,attr"`
	Scale          int               `xml:"scale,attr,omitempty"`
	ColorID        int               `xml:"colorId,attr,omitempty"`
	ShowPageBreaks bool              `xml:"showPageBreaks,attr,omitempty"`
	ShowFormulas   bool              `xml:"showFormulas,attr,omitempty"`
	ShowGridLines  bool              `xml:"showGridLines,attr,omitempty"`
	ShowRowCol     bool              `xml:"showRowCol,attr,omitempty"`
	OutlineSymbols bool              `xml:"outlineSymbols,attr,omitempty"`
	ZeroValues     bool              `xml:"zeroValues,attr,omitempty"`
	FitToPage      bool              `xml:"fitToPage,attr,omitempty"`
	PrintArea      bool              `xml:"printArea,attr,omitempty"`
	Filter         bool              `xml:"filter,attr,omitempty"`
	ShowAutoFilter bool              `xml:"showAutoFilter,attr,omitempty"`
	HiddenRows     bool              `xml:"hiddenRows,attr,omitempty"`
	HiddenColumns  bool              `xml:"hiddenColumns,attr,omitempty"`
	State          string            `xml:"state,attr,omitempty"`
	FilterUnique   bool              `xml:"filterUnique,attr,omitempty"`
	View           string            `xml:"view,attr,omitempty"`
	ShowRuler      bool              `xml:"showRuler,attr,omitempty"`
	TopLeftCell    string            `xml:"topLeftCell,attr,omitempty"`
}

// xlsxMergeCell directly maps the mergeCell element. A single merged cell.
type xlsxMergeCell struct {
	Ref  string `xml:"ref,attr,omitempty"`
	rect []int
}

// xlsxMergeCells directly maps the mergeCells element. This collection
// expresses all the merged cells in the sheet.
type xlsxMergeCells struct {
	XMLName xml.Name         `xml:"mergeCells"`
	Count   int              `xml:"count,attr,omitempty"`
	Cells   []*xlsxMergeCell `xml:"mergeCell,omitempty"`
}

// xlsxDataValidations expresses all data validation information for cells in a
// sheet which have data validation features applied.
type xlsxDataValidations struct {
	XMLName        xml.Name          `xml:"dataValidations"`
	Count          int               `xml:"count,attr,omitempty"`
	DisablePrompts bool              `xml:"disablePrompts,attr,omitempty"`
	XWindow        int               `xml:"xWindow,attr,omitempty"`
	YWindow        int               `xml:"yWindow,attr,omitempty"`
	DataValidation []*DataValidation `xml:"dataValidation"`
}

// DataValidation directly maps the a single item of data validation defined
// on a range of the worksheet.
type DataValidation struct {
	AllowBlank       bool    `xml:"allowBlank,attr"`
	Error            *string `xml:"error,attr"`
	ErrorStyle       *string `xml:"errorStyle,attr"`
	ErrorTitle       *string `xml:"errorTitle,attr"`
	Operator         string  `xml:"operator,attr,omitempty"`
	Prompt           *string `xml:"prompt,attr"`
	PromptTitle      *string `xml:"promptTitle,attr"`
	ShowDropDown     bool    `xml:"showDropDown,attr,omitempty"`
	ShowErrorMessage bool    `xml:"showErrorMessage,attr,omitempty"`
	ShowInputMessage bool    `xml:"showInputMessage,attr,omitempty"`
	Sqref            string  `xml:"sqref,attr"`
	Type             string  `xml:"type,attr,omitempty"`
	Formula1         string  `xml:",innerxml"`
	Formula2         string  `xml:",innerxml"`
}

// xlsxC collection represents a cell in the worksheet. Information about the
// cell's location (reference), value, data type, formatting, and formula is
// expressed here.
//
// This simple type is restricted to the values listed in the following table:
//
//      Enumeration Value         | Description
//     ---------------------------+---------------------------------
//      b (Boolean)               | Cell containing a boolean.
//      d (Date)                  | Cell contains a date in the ISO 8601 format.
//      e (Error)                 | Cell containing an error.
//      inlineStr (Inline String) | Cell containing an (inline) rich string, i.e., one not in the shared string table. If this cell type is used, then the cell value is in the is element rather than the v element in the cell (c element).
//      n (Number)                | Cell containing a number.
//      s (Shared String)         | Cell containing a shared string.
//      str (String)              | Cell containing a formula string.
//
type xlsxC struct {
	XMLName  xml.Name `xml:"c"`
	XMLSpace xml.Attr `xml:"space,attr,omitempty"`
	R        string   `xml:"r,attr,omitempty"` // Cell ID, e.g. A1
	S        int      `xml:"s,attr,omitempty"` // Style reference.
	// Str string `xml:"str,attr,omitempty"` // Style reference.
	T  string  `xml:"t,attr,omitempty"` // Type.
	F  *xlsxF  `xml:"f,omitempty"`      // Formula
	V  string  `xml:"v,omitempty"`      // Value
	IS *xlsxSI `xml:"is"`
}

// xlsxF represents a formula for the cell. The formula expression is
// contained in the character node of this element.
type xlsxF struct {
	Content string `xml:",chardata"`
	T       string `xml:"t,attr,omitempty"` // Formula type
	Aca     bool   `xml:"aca,attr,omitempty"`
	Ref     string `xml:"ref,attr,omitempty"` // Shared formula ref
	Dt2D    bool   `xml:"dt2D,attr,omitempty"`
	Dtr     bool   `xml:"dtr,attr,omitempty"`
	Del1    bool   `xml:"del1,attr,omitempty"`
	Del2    bool   `xml:"del2,attr,omitempty"`
	R1      string `xml:"r1,attr,omitempty"`
	R2      string `xml:"r2,attr,omitempty"`
	Ca      bool   `xml:"ca,attr,omitempty"`
	Si      *int   `xml:"si,attr"` // Shared formula index
	Bx      bool   `xml:"bx,attr,omitempty"`
}

// xlsxSheetProtection collection expresses the sheet protection options to
// enforce when the sheet is protected.
type xlsxSheetProtection struct {
	XMLName             xml.Name `xml:"sheetProtection"`
	AlgorithmName       string   `xml:"algorithmName,attr,omitempty"`
	Password            string   `xml:"password,attr,omitempty"`
	HashValue           string   `xml:"hashValue,attr,omitempty"`
	SaltValue           string   `xml:"saltValue,attr,omitempty"`
	SpinCount           int      `xml:"spinCount,attr,omitempty"`
	Sheet               bool     `xml:"sheet,attr"`
	Objects             bool     `xml:"objects,attr"`
	Scenarios           bool     `xml:"scenarios,attr"`
	FormatCells         bool     `xml:"formatCells,attr"`
	FormatColumns       bool     `xml:"formatColumns,attr"`
	FormatRows          bool     `xml:"formatRows,attr"`
	InsertColumns       bool     `xml:"insertColumns,attr"`
	InsertRows          bool     `xml:"insertRows,attr"`
	InsertHyperlinks    bool     `xml:"insertHyperlinks,attr"`
	DeleteColumns       bool     `xml:"deleteColumns,attr"`
	DeleteRows          bool     `xml:"deleteRows,attr"`
	SelectLockedCells   bool     `xml:"selectLockedCells,attr"`
	Sort                bool     `xml:"sort,attr"`
	AutoFilter          bool     `xml:"autoFilter,attr"`
	PivotTables         bool     `xml:"pivotTables,attr"`
	SelectUnlockedCells bool     `xml:"selectUnlockedCells,attr"`
}

// xlsxPhoneticPr (Phonetic Properties) represents a collection of phonetic
// properties that affect the display of phonetic text for this String Item
// (si). Phonetic text is used to give hints as to the pronunciation of an East
// Asian language, and the hints are displayed as text within the spreadsheet
// cells across the top portion of the cell. Since the phonetic hints are text,
// every phonetic hint is expressed as a phonetic run (rPh), and these
// properties specify how to display that phonetic run.
type xlsxPhoneticPr struct {
	XMLName   xml.Name `xml:"phoneticPr"`
	Alignment string   `xml:"alignment,attr,omitempty"`
	FontID    *int     `xml:"fontId,attr"`
	Type      string   `xml:"type,attr,omitempty"`
}

// A Conditional Format is a format, such as cell shading or font color, that a
// spreadsheet application can automatically apply to cells if a specified
// condition is true. This collection expresses conditional formatting rules
// applied to a particular cell or range.
type xlsxConditionalFormatting struct {
	XMLName xml.Name      `xml:"conditionalFormatting"`
	Pivot   bool          `xml:"pivot,attr,omitempty"`
	SQRef   string        `xml:"sqref,attr,omitempty"`
	CfRule  []*xlsxCfRule `xml:"cfRule"`
}

// xlsxCfRule (Conditional Formatting Rule) represents a description of a
// conditional formatting rule.
type xlsxCfRule struct {
	Type         string          `xml:"type,attr,omitempty"`
	DxfID        *int            `xml:"dxfId,attr"`
	Priority     int             `xml:"priority,attr,omitempty"`
	StopIfTrue   bool            `xml:"stopIfTrue,attr,omitempty"`
	AboveAverage *bool           `xml:"aboveAverage,attr"`
	Percent      bool            `xml:"percent,attr,omitempty"`
	Bottom       bool            `xml:"bottom,attr,omitempty"`
	Operator     string          `xml:"operator,attr,omitempty"`
	Text         string          `xml:"text,attr,omitempty"`
	TimePeriod   string          `xml:"timePeriod,attr,omitempty"`
	Rank         int             `xml:"rank,attr,omitempty"`
	StdDev       int             `xml:"stdDev,attr,omitempty"`
	EqualAverage bool            `xml:"equalAverage,attr,omitempty"`
	Formula      []string        `xml:"formula,omitempty"`
	ColorScale   *xlsxColorScale `xml:"colorScale"`
	DataBar      *xlsxDataBar    `xml:"dataBar"`
	IconSet      *xlsxIconSet    `xml:"iconSet"`
	ExtLst       *xlsxExtLst     `xml:"extLst"`
}

// xlsxColorScale (Color Scale) describes a gradated color scale in this
// conditional formatting rule.
type xlsxColorScale struct {
	Cfvo  []*xlsxCfvo  `xml:"cfvo"`
	Color []*xlsxColor `xml:"color"`
}

// dataBar (Data Bar) describes a data bar conditional formatting rule.
type xlsxDataBar struct {
	MaxLength int          `xml:"maxLength,attr,omitempty"`
	MinLength int          `xml:"minLength,attr,omitempty"`
	ShowValue bool         `xml:"showValue,attr,omitempty"`
	Cfvo      []*xlsxCfvo  `xml:"cfvo"`
	Color     []*xlsxColor `xml:"color"`
}

// xlsxIconSet (Icon Set) describes an icon set conditional formatting rule.
type xlsxIconSet struct {
	Cfvo      []*xlsxCfvo `xml:"cfvo"`
	IconSet   string      `xml:"iconSet,attr,omitempty"`
	ShowValue bool        `xml:"showValue,attr,omitempty"`
	Percent   bool        `xml:"percent,attr,omitempty"`
	Reverse   bool        `xml:"reverse,attr,omitempty"`
}

// cfvo (Conditional Format Value Object) describes the values of the
// interpolation points in a gradient scale.
type xlsxCfvo struct {
	Gte    bool        `xml:"gte,attr,omitempty"`
	Type   string      `xml:"type,attr,omitempty"`
	Val    string      `xml:"val,attr,omitempty"`
	ExtLst *xlsxExtLst `xml:"extLst"`
}

// xlsxHyperlinks directly maps the hyperlinks element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - A hyperlink can
// be stored in a package as a relationship. Hyperlinks shall be identified by
// containing a target which specifies the destination of the given hyperlink.
type xlsxHyperlinks struct {
	XMLName   xml.Name        `xml:"hyperlinks"`
	Hyperlink []xlsxHyperlink `xml:"hyperlink"`
}

// xlsxHyperlink directly maps the hyperlink element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxHyperlink struct {
	Ref      string `xml:"ref,attr"`
	Location string `xml:"location,attr,omitempty"`
	Display  string `xml:"display,attr,omitempty"`
	Tooltip  string `xml:"tooltip,attr,omitempty"`
	RID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// xlsxTableParts directly maps the tableParts element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - The table element
// has several attributes applied to identify the table and the data range it
// covers. The table id attribute needs to be unique across all table parts, the
// same goes for the name and displayName. The displayName has the further
// restriction that it must be unique across all defined names in the workbook.
// Later on we will see that you can define names for many elements, such as
// cells or formulas. The name value is used for the object model in Microsoft
// Office Excel. The displayName is used for references in formulas. The ref
// attribute is used to identify the cell range that the table covers. This
// includes not only the table data, but also the table header containing column
// names.
// To add columns to your table you add new tableColumn elements to the
// tableColumns container. Similar to the shared string table the collection
// keeps a count attribute identifying the number of columns. Besides the table
// definition in the table part there is also the need to identify which tables
// are displayed in the worksheet. The worksheet part has a separate element
// tableParts to store this information. Each table part is referenced through
// the relationship ID and again a count of the number of table parts is
// maintained. The following markup sample is taken from the documents
// accompanying this book. The sheet data element has been removed to reduce the
// size of the sample. To reference the table, just add the tableParts element,
// of course after having created and stored the table part. For example:
//
//    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//        ...
//        <tableParts count="1">
// 		      <tablePart r:id="rId1" />
//        </tableParts>
//    </worksheet>
//
type xlsxTableParts struct {
	XMLName    xml.Name         `xml:"tableParts"`
	Count      int              `xml:"count,attr,omitempty"`
	TableParts []*xlsxTablePart `xml:"tablePart"`
}

// xlsxTablePart directly maps the tablePart element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxTablePart struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// xlsxPicture directly maps the picture element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - Background sheet
// image. For example:
//
//    <picture r:id="rId1"/>
//
type xlsxPicture struct {
	XMLName xml.Name `xml:"picture"`
	RID     string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// xlsxLegacyDrawing directly maps the legacyDrawing element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main - A comment is a
// rich text note that is attached to, and associated with, a cell, separate
// from other cell content. Comment content is stored separate from the cell,
// and is displayed in a drawing object (like a text box) that is separate from,
// but associated with, a cell. Comments are used as reminders, such as noting
// how a complex formula works, or to provide feedback to other users. Comments
// can also be used to explain assumptions made in a formula or to call out
// something special about the cell.
type xlsxLegacyDrawing struct {
	XMLName xml.Name `xml:"legacyDrawing"`
	RID     string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// xlsxLegacyDrawingHF specifies the explicit relationship to the part
// containing the VML defining pictures rendered in the header / footer of the
// sheet.
type xlsxLegacyDrawingHF struct {
	XMLName xml.Name `xml:"legacyDrawingHF"`
	RID     string   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

type xlsxInnerXML struct {
	Content string `xml:",innerxml"`
}

// xlsxWorksheetExt directly maps the ext element in the worksheet.
type xlsxWorksheetExt struct {
	XMLName xml.Name `xml:"ext"`
	URI     string   `xml:"uri,attr"`
	Content string   `xml:",innerxml"`
}

// decodeWorksheetExt directly maps the ext element.
type decodeWorksheetExt struct {
	XMLName xml.Name            `xml:"extLst"`
	Ext     []*xlsxWorksheetExt `xml:"ext"`
}

// decodeX14SparklineGroups directly maps the sparklineGroups element.
type decodeX14SparklineGroups struct {
	XMLName xml.Name `xml:"sparklineGroups"`
	XMLNSXM string   `xml:"xmlns:xm,attr"`
	Content string   `xml:",innerxml"`
}

// xlsxX14SparklineGroups directly maps the sparklineGroups element.
type xlsxX14SparklineGroups struct {
	XMLName         xml.Name                 `xml:"x14:sparklineGroups"`
	XMLNSXM         string                   `xml:"xmlns:xm,attr"`
	SparklineGroups []*xlsxX14SparklineGroup `xml:"x14:sparklineGroup"`
	Content         string                   `xml:",innerxml"`
}

// xlsxX14SparklineGroup directly maps the sparklineGroup element.
type xlsxX14SparklineGroup struct {
	XMLName             xml.Name          `xml:"x14:sparklineGroup"`
	ManualMax           int               `xml:"manualMax,attr,omitempty"`
	ManualMin           int               `xml:"manualMin,attr,omitempty"`
	LineWeight          float64           `xml:"lineWeight,attr,omitempty"`
	Type                string            `xml:"type,attr,omitempty"`
	DateAxis            bool              `xml:"dateAxis,attr,omitempty"`
	DisplayEmptyCellsAs string            `xml:"displayEmptyCellsAs,attr,omitempty"`
	Markers             bool              `xml:"markers,attr,omitempty"`
	High                bool              `xml:"high,attr,omitempty"`
	Low                 bool              `xml:"low,attr,omitempty"`
	First               bool              `xml:"first,attr,omitempty"`
	Last                bool              `xml:"last,attr,omitempty"`
	Negative            bool              `xml:"negative,attr,omitempty"`
	DisplayXAxis        bool              `xml:"displayXAxis,attr,omitempty"`
	DisplayHidden       bool              `xml:"displayHidden,attr,omitempty"`
	MinAxisType         string            `xml:"minAxisType,attr,omitempty"`
	MaxAxisType         string            `xml:"maxAxisType,attr,omitempty"`
	RightToLeft         bool              `xml:"rightToLeft,attr,omitempty"`
	ColorSeries         *xlsxTabColor     `xml:"x14:colorSeries"`
	ColorNegative       *xlsxTabColor     `xml:"x14:colorNegative"`
	ColorAxis           *xlsxColor        `xml:"x14:colorAxis"`
	ColorMarkers        *xlsxTabColor     `xml:"x14:colorMarkers"`
	ColorFirst          *xlsxTabColor     `xml:"x14:colorFirst"`
	ColorLast           *xlsxTabColor     `xml:"x14:colorLast"`
	ColorHigh           *xlsxTabColor     `xml:"x14:colorHigh"`
	ColorLow            *xlsxTabColor     `xml:"x14:colorLow"`
	Sparklines          xlsxX14Sparklines `xml:"x14:sparklines"`
}

// xlsxX14Sparklines directly maps the sparklines element.
type xlsxX14Sparklines struct {
	Sparkline []*xlsxX14Sparkline `xml:"x14:sparkline"`
}

// xlsxX14Sparkline directly maps the sparkline element.
type xlsxX14Sparkline struct {
	F     string `xml:"xm:f"`
	Sqref string `xml:"xm:sqref"`
}

// SparklineOption directly maps the settings of the sparkline.
type SparklineOption struct {
	Location      []string
	Range         []string
	Max           int
	CustMax       int
	Min           int
	CustMin       int
	Type          string
	Weight        float64
	DateAxis      bool
	Markers       bool
	High          bool
	Low           bool
	First         bool
	Last          bool
	Negative      bool
	Axis          bool
	Hidden        bool
	Reverse       bool
	Style         int
	SeriesColor   string
	NegativeColor string
	MarkersColor  string
	FirstColor    string
	LastColor     string
	HightColor    string
	LowColor      string
	EmptyCells    string
}

// formatPanes directly maps the settings of the panes.
type formatPanes struct {
	Freeze      bool   `json:"freeze"`
	Split       bool   `json:"split"`
	XSplit      int    `json:"x_split"`
	YSplit      int    `json:"y_split"`
	TopLeftCell string `json:"top_left_cell"`
	ActivePane  string `json:"active_pane"`
	Panes       []struct {
		SQRef      string `json:"sqref"`
		ActiveCell string `json:"active_cell"`
		Pane       string `json:"pane"`
	} `json:"panes"`
}

// formatConditional directly maps the conditional format settings of the cells.
type formatConditional struct {
	Type         string `json:"type"`
	AboveAverage bool   `json:"above_average"`
	Percent      bool   `json:"percent"`
	Format       int    `json:"format"`
	Criteria     string `json:"criteria"`
	Value        string `json:"value,omitempty"`
	Minimum      string `json:"minimum,omitempty"`
	Maximum      string `json:"maximum,omitempty"`
	MinType      string `json:"min_type,omitempty"`
	MidType      string `json:"mid_type,omitempty"`
	MaxType      string `json:"max_type,omitempty"`
	MinValue     string `json:"min_value,omitempty"`
	MidValue     string `json:"mid_value,omitempty"`
	MaxValue     string `json:"max_value,omitempty"`
	MinColor     string `json:"min_color,omitempty"`
	MidColor     string `json:"mid_color,omitempty"`
	MaxColor     string `json:"max_color,omitempty"`
	MinLength    string `json:"min_length,omitempty"`
	MaxLength    string `json:"max_length,omitempty"`
	MultiRange   string `json:"multi_range,omitempty"`
	BarColor     string `json:"bar_color,omitempty"`
}

// FormatSheetProtection directly maps the settings of worksheet protection.
type FormatSheetProtection struct {
	AutoFilter          bool
	DeleteColumns       bool
	DeleteRows          bool
	EditObjects         bool
	EditScenarios       bool
	FormatCells         bool
	FormatColumns       bool
	FormatRows          bool
	InsertColumns       bool
	InsertHyperlinks    bool
	InsertRows          bool
	Password            string
	PivotTables         bool
	SelectLockedCells   bool
	SelectUnlockedCells bool
	Sort                bool
}

// FormatHeaderFooter directly maps the settings of header and footer.
type FormatHeaderFooter struct {
	AlignWithMargins bool
	DifferentFirst   bool
	DifferentOddEven bool
	ScaleWithDoc     bool
	OddHeader        string
	OddFooter        string
	EvenHeader       string
	EvenFooter       string
	FirstFooter      string
	FirstHeader      string
}

// FormatPageMargins directly maps the settings of page margins
type FormatPageMargins struct {
	Bottom string
	Footer string
	Header string
	Left   string
	Right  string
	Top    string
}
