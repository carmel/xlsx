package xlsx

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// PivotTableOption directly maps the format settings of the pivot table.
//
// PivotTableStyleName: The built-in pivot table style names
//
//    PivotStyleLight1 - PivotStyleLight28
//    PivotStyleMedium1 - PivotStyleMedium28
//    PivotStyleDark1 - PivotStyleDark28
//
type PivotTableOption struct {
	pivotTableSheetName string
	DataRange           string
	PivotTableRange     string
	Rows                []PivotTableField
	Columns             []PivotTableField
	Data                []PivotTableField
	Filter              []PivotTableField
	RowGrandTotals      bool
	ColGrandTotals      bool
	ShowDrill           bool
	UseAutoFormatting   bool
	PageOverThenDown    bool
	MergeItem           bool
	CompactData         bool
	ShowError           bool
	ShowRowHeaders      bool
	ShowColHeaders      bool
	ShowRowStripes      bool
	ShowColStripes      bool
	ShowLastColumn      bool
	PivotTableStyleName string
}

// PivotTableField directly maps the field settings of the pivot table.
// Subtotal specifies the aggregation function that applies to this data
// field. The default value is sum. The possible values for this attribute
// are:
//
//     Average
//     Count
//     CountNums
//     Max
//     Min
//     Product
//     StdDev
//     StdDevp
//     Sum
//     Var
//     Varp
//
// Name specifies the name of the data field. Maximum 255 characters
// are allowed in data field name, excess characters will be truncated.
type PivotTableField struct {
	Compact         bool
	Data            string
	Name            string
	Outline         bool
	Subtotal        string
	DefaultSubtotal bool
}

// AddPivotTable provides the method to add pivot table by given pivot table
// options. Note that the same fields can not in Columns, Rows and Filter
// fields at the same time.
//
// For example, create a pivot table on the Sheet1!$G$2:$M$34 area with the
// region Sheet1!$A$1:$E$31 as the data source, summarize by sum for sales:
//
//    package main
//
//    import (
//        "fmt"
//        "math/rand"
//
//        "github.com/carmel/xlsx"
//    )
//
//    func main() {
//        f := xlsx.NewFile()
//        // Create some data in a sheet
//        month := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
//        year := []int{2017, 2018, 2019}
//        types := []string{"Meat", "Dairy", "Beverages", "Produce"}
//        region := []string{"East", "West", "North", "South"}
//        f.SetSheetRow("Sheet1", "A1", &[]string{"Month", "Year", "Type", "Sales", "Region"})
//        for i := 0; i < 30; i++ {
//            f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+2), month[rand.Intn(12)])
//            f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+2), year[rand.Intn(3)])
//            f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i+2), types[rand.Intn(4)])
//            f.SetCellValue("Sheet1", fmt.Sprintf("D%d", i+2), rand.Intn(5000))
//            f.SetCellValue("Sheet1", fmt.Sprintf("E%d", i+2), region[rand.Intn(4)])
//        }
//        if err := f.AddPivotTable(&xlsx.PivotTableOption{
//            DataRange:       "Sheet1!$A$1:$E$31",
//            PivotTableRange: "Sheet1!$G$2:$M$34",
//            Rows:            []xlsx.PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
//            Filter:          []xlsx.PivotTableField{{Data: "Region"}},
//            Columns:         []xlsx.PivotTableField{{Data: "Type", DefaultSubtotal: true}},
//            Data:            []xlsx.PivotTableField{{Data: "Sales", Name: "Summarize", Subtotal: "Sum"}},
//            RowGrandTotals:  true,
//            ColGrandTotals:  true,
//            ShowDrill:       true,
//            ShowRowHeaders:  true,
//            ShowColHeaders:  true,
//            ShowLastColumn:  true,
//        }); err != nil {
//            fmt.Println(err)
//        }
//        if err := f.SaveAs("Book1.xlsx"); err != nil {
//            fmt.Println(err)
//        }
//    }
//
func (f *File) AddPivotTable(opt *PivotTableOption) error {
	// parameter validation
	dataSheet, pivotTableSheetPath, err := f.parseFormatPivotTableSet(opt)
	if err != nil {
		return err
	}

	pivotTableID := f.countPivotTables() + 1
	pivotCacheID := f.countPivotCache() + 1

	sheetRelationshipsPivotTableXML := "../pivotTables/pivotTable" + strconv.Itoa(pivotTableID) + ".xml"
	pivotTableXML := strings.Replace(sheetRelationshipsPivotTableXML, "..", "xl", -1)
	pivotCacheXML := "xl/pivotCache/pivotCacheDefinition" + strconv.Itoa(pivotCacheID) + ".xml"
	err = f.addPivotCache(pivotCacheID, pivotCacheXML, opt, dataSheet)
	if err != nil {
		return err
	}

	// workbook pivot cache
	workBookPivotCacheRID := f.addRels(f.getWorkbookRelsPath(), SourceRelationshipPivotCache, fmt.Sprintf("/xl/pivotCache/pivotCacheDefinition%d.xml", pivotCacheID), "")
	cacheID := f.addWorkbookPivotCache(workBookPivotCacheRID)

	pivotCacheRels := "xl/pivotTables/_rels/pivotTable" + strconv.Itoa(pivotTableID) + ".xml.rels"
	// rId not used
	_ = f.addRels(pivotCacheRels, SourceRelationshipPivotCache, fmt.Sprintf("../pivotCache/pivotCacheDefinition%d.xml", pivotCacheID), "")
	err = f.addPivotTable(cacheID, pivotTableID, pivotTableXML, opt)
	if err != nil {
		return err
	}
	pivotTableSheetRels := "xl/worksheets/_rels/" + strings.TrimPrefix(pivotTableSheetPath, "xl/worksheets/") + ".rels"
	f.addRels(pivotTableSheetRels, SourceRelationshipPivotTable, sheetRelationshipsPivotTableXML, "")
	f.addContentTypePart(pivotTableID, "pivotTable")
	f.addContentTypePart(pivotCacheID, "pivotCache")

	return nil
}

// parseFormatPivotTableSet provides a function to validate pivot table
// properties.
func (f *File) parseFormatPivotTableSet(opt *PivotTableOption) (*xlsxWorksheet, string, error) {
	if opt == nil {
		return nil, "", ErrParameterRequired
	}
	pivotTableSheetName, _, err := f.adjustRange(opt.PivotTableRange)
	if err != nil {
		return nil, "", fmt.Errorf("parameter 'PivotTableRange' parsing error: %s", err.Error())
	}
	opt.pivotTableSheetName = pivotTableSheetName
	dataRange := f.getDefinedNameRefTo(opt.DataRange, pivotTableSheetName)
	if dataRange == "" {
		dataRange = opt.DataRange
	}
	dataSheetName, _, err := f.adjustRange(dataRange)
	if err != nil {
		return nil, "", fmt.Errorf("parameter 'DataRange' parsing error: %s", err.Error())
	}
	dataSheet, err := f.workSheetReader(dataSheetName)
	if err != nil {
		return dataSheet, "", err
	}
	pivotTableSheetPath, ok := f.sheetMap[trimSheetName(pivotTableSheetName)]
	if !ok {
		return dataSheet, pivotTableSheetPath, fmt.Errorf("sheet %s is not exist", pivotTableSheetName)
	}
	return dataSheet, pivotTableSheetPath, err
}

// adjustRange adjust range, for example: adjust Sheet1!$E$31:$A$1 to Sheet1!$A$1:$E$31
func (f *File) adjustRange(rangeStr string) (string, []int, error) {
	if len(rangeStr) < 1 {
		return "", []int{}, ErrParameterRequired
	}
	rng := strings.Split(rangeStr, "!")
	if len(rng) != 2 {
		return "", []int{}, ErrParameterInvalid
	}
	trimRng := strings.Replace(rng[1], "$", "", -1)
	coordinates, err := areaRefToCoordinates(trimRng)
	if err != nil {
		return rng[0], []int{}, err
	}
	x1, y1, x2, y2 := coordinates[0], coordinates[1], coordinates[2], coordinates[3]
	if x1 == x2 && y1 == y2 {
		return rng[0], []int{}, ErrParameterInvalid
	}

	// Correct the coordinate area, such correct C1:B3 to B1:C3.
	if x2 < x1 {
		x1, x2 = x2, x1
	}

	if y2 < y1 {
		y1, y2 = y2, y1
	}
	return rng[0], []int{x1, y1, x2, y2}, nil
}

// getPivotFieldsOrder provides a function to get order list of pivot table
// fields.
func (f *File) getPivotFieldsOrder(opt *PivotTableOption) ([]string, error) {
	order := []string{}
	dataRange := f.getDefinedNameRefTo(opt.DataRange, opt.pivotTableSheetName)
	if dataRange == "" {
		dataRange = opt.DataRange
	}
	dataSheet, coordinates, err := f.adjustRange(dataRange)
	if err != nil {
		return order, fmt.Errorf("parameter 'DataRange' parsing error: %s", err.Error())
	}
	for col := coordinates[0]; col <= coordinates[2]; col++ {
		coordinate, _ := CoordinatesToCellName(col, coordinates[1])
		name, err := f.GetCellValue(dataSheet, coordinate)
		if err != nil {
			return order, err
		}
		order = append(order, name)
	}
	return order, nil
}

// addPivotCache provides a function to create a pivot cache by given properties.
func (f *File) addPivotCache(pivotCacheID int, pivotCacheXML string, opt *PivotTableOption, ws *xlsxWorksheet) error {
	// validate data range
	definedNameRef := true
	dataRange := f.getDefinedNameRefTo(opt.DataRange, opt.pivotTableSheetName)
	if dataRange == "" {
		definedNameRef = false
		dataRange = opt.DataRange
	}
	dataSheet, coordinates, err := f.adjustRange(dataRange)
	if err != nil {
		return fmt.Errorf("parameter 'DataRange' parsing error: %s", err.Error())
	}
	// data range has been checked
	order, _ := f.getPivotFieldsOrder(opt)
	hcell, _ := CoordinatesToCellName(coordinates[0], coordinates[1])
	vcell, _ := CoordinatesToCellName(coordinates[2], coordinates[3])
	pc := xlsxPivotCacheDefinition{
		SaveData:              false,
		RefreshOnLoad:         true,
		CreatedVersion:        pivotTableVersion,
		RefreshedVersion:      pivotTableVersion,
		MinRefreshableVersion: pivotTableVersion,
		CacheSource: &xlsxCacheSource{
			Type: "worksheet",
			WorksheetSource: &xlsxWorksheetSource{
				Ref:   hcell + ":" + vcell,
				Sheet: dataSheet,
			},
		},
		CacheFields: &xlsxCacheFields{},
	}
	if definedNameRef {
		pc.CacheSource.WorksheetSource = &xlsxWorksheetSource{Name: opt.DataRange}
	}
	for _, name := range order {
		rowOptions, rowOk := f.getPivotTableFieldOptions(name, opt.Rows)
		columnOptions, colOk := f.getPivotTableFieldOptions(name, opt.Columns)
		sharedItems := xlsxSharedItems{
			Count: 0,
		}
		s := xlsxString{}
		if (rowOk && !rowOptions.DefaultSubtotal) || (colOk && !columnOptions.DefaultSubtotal) {
			s = xlsxString{
				V: "",
			}
			sharedItems.Count++
			sharedItems.S = &s
		}

		pc.CacheFields.CacheField = append(pc.CacheFields.CacheField, &xlsxCacheField{
			Name:        name,
			SharedItems: &sharedItems,
		})
	}
	pc.CacheFields.Count = len(pc.CacheFields.CacheField)
	pivotCache, err := xml.Marshal(pc)
	f.saveFileList(pivotCacheXML, pivotCache)
	return err
}

// addPivotTable provides a function to create a pivot table by given pivot
// table ID and properties.
func (f *File) addPivotTable(cacheID, pivotTableID int, pivotTableXML string, opt *PivotTableOption) error {
	// validate pivot table range
	_, coordinates, err := f.adjustRange(opt.PivotTableRange)
	if err != nil {
		return fmt.Errorf("parameter 'PivotTableRange' parsing error: %s", err.Error())
	}

	hcell, _ := CoordinatesToCellName(coordinates[0], coordinates[1])
	vcell, _ := CoordinatesToCellName(coordinates[2], coordinates[3])

	pivotTableStyle := func() string {
		if opt.PivotTableStyleName == "" {
			return "PivotStyleLight16"
		}
		return opt.PivotTableStyleName
	}
	pt := xlsxPivotTableDefinition{
		Name:                  fmt.Sprintf("Pivot Table%d", pivotTableID),
		CacheID:               cacheID,
		RowGrandTotals:        &opt.RowGrandTotals,
		ColGrandTotals:        &opt.ColGrandTotals,
		UpdatedVersion:        pivotTableVersion,
		MinRefreshableVersion: pivotTableVersion,
		ShowDrill:             &opt.ShowDrill,
		UseAutoFormatting:     &opt.UseAutoFormatting,
		PageOverThenDown:      &opt.PageOverThenDown,
		MergeItem:             &opt.MergeItem,
		CreatedVersion:        pivotTableVersion,
		CompactData:           &opt.CompactData,
		ShowError:             &opt.ShowError,
		DataCaption:           "Values",
		Location: &xlsxLocation{
			Ref:            hcell + ":" + vcell,
			FirstDataCol:   1,
			FirstDataRow:   1,
			FirstHeaderRow: 1,
		},
		PivotFields: &xlsxPivotFields{},
		RowItems: &xlsxRowItems{
			Count: 1,
			I: []*xlsxI{
				{
					[]*xlsxX{{}, {}},
				},
			},
		},
		ColItems: &xlsxColItems{
			Count: 1,
			I:     []*xlsxI{{}},
		},
		PivotTableStyleInfo: &xlsxPivotTableStyleInfo{
			Name:           pivotTableStyle(),
			ShowRowHeaders: opt.ShowRowHeaders,
			ShowColHeaders: opt.ShowColHeaders,
			ShowRowStripes: opt.ShowRowStripes,
			ShowColStripes: opt.ShowColStripes,
			ShowLastColumn: opt.ShowLastColumn,
		},
	}

	// pivot fields
	_ = f.addPivotFields(&pt, opt)

	// count pivot fields
	pt.PivotFields.Count = len(pt.PivotFields.PivotField)

	// data range has been checked
	_ = f.addPivotRowFields(&pt, opt)
	_ = f.addPivotColFields(&pt, opt)
	_ = f.addPivotPageFields(&pt, opt)
	_ = f.addPivotDataFields(&pt, opt)

	pivotTable, err := xml.Marshal(pt)
	f.saveFileList(pivotTableXML, pivotTable)
	return err
}

// addPivotRowFields provides a method to add row fields for pivot table by
// given pivot table options.
func (f *File) addPivotRowFields(pt *xlsxPivotTableDefinition, opt *PivotTableOption) error {
	// row fields
	rowFieldsIndex, err := f.getPivotFieldsIndex(opt.Rows, opt)
	if err != nil {
		return err
	}
	for _, fieldIdx := range rowFieldsIndex {
		if pt.RowFields == nil {
			pt.RowFields = &xlsxRowFields{}
		}
		pt.RowFields.Field = append(pt.RowFields.Field, &xlsxField{
			X: fieldIdx,
		})
	}

	// count row fields
	if pt.RowFields != nil {
		pt.RowFields.Count = len(pt.RowFields.Field)
	}
	return err
}

// addPivotPageFields provides a method to add page fields for pivot table by
// given pivot table options.
func (f *File) addPivotPageFields(pt *xlsxPivotTableDefinition, opt *PivotTableOption) error {
	// page fields
	pageFieldsIndex, err := f.getPivotFieldsIndex(opt.Filter, opt)
	if err != nil {
		return err
	}
	pageFieldsName := f.getPivotTableFieldsName(opt.Filter)
	for idx, pageField := range pageFieldsIndex {
		if pt.PageFields == nil {
			pt.PageFields = &xlsxPageFields{}
		}
		pt.PageFields.PageField = append(pt.PageFields.PageField, &xlsxPageField{
			Name: pageFieldsName[idx],
			Fld:  pageField,
		})
	}

	// count page fields
	if pt.PageFields != nil {
		pt.PageFields.Count = len(pt.PageFields.PageField)
	}
	return err
}

// addPivotDataFields provides a method to add data fields for pivot table by
// given pivot table options.
func (f *File) addPivotDataFields(pt *xlsxPivotTableDefinition, opt *PivotTableOption) error {
	// data fields
	dataFieldsIndex, err := f.getPivotFieldsIndex(opt.Data, opt)
	if err != nil {
		return err
	}
	dataFieldsSubtotals := f.getPivotTableFieldsSubtotal(opt.Data)
	dataFieldsName := f.getPivotTableFieldsName(opt.Data)
	for idx, dataField := range dataFieldsIndex {
		if pt.DataFields == nil {
			pt.DataFields = &xlsxDataFields{}
		}
		pt.DataFields.DataField = append(pt.DataFields.DataField, &xlsxDataField{
			Name:     dataFieldsName[idx],
			Fld:      dataField,
			Subtotal: dataFieldsSubtotals[idx],
		})
	}

	// count data fields
	if pt.DataFields != nil {
		pt.DataFields.Count = len(pt.DataFields.DataField)
	}
	return err
}

// inPivotTableField provides a method to check if an element is present in
// pivot table fields list, and return the index of its location, otherwise
// return -1.
func inPivotTableField(a []PivotTableField, x string) int {
	for idx, n := range a {
		if x == n.Data {
			return idx
		}
	}
	return -1
}

// addPivotColFields create pivot column fields by given pivot table
// definition and option.
func (f *File) addPivotColFields(pt *xlsxPivotTableDefinition, opt *PivotTableOption) error {
	if len(opt.Columns) == 0 {
		if len(opt.Data) <= 1 {
			return nil
		}
		pt.ColFields = &xlsxColFields{}
		// in order to create pivot table in case there is no input from Columns
		pt.ColFields.Count = 1
		pt.ColFields.Field = append(pt.ColFields.Field, &xlsxField{
			X: -2,
		})
		return nil
	}

	pt.ColFields = &xlsxColFields{}

	// col fields
	colFieldsIndex, err := f.getPivotFieldsIndex(opt.Columns, opt)
	if err != nil {
		return err
	}
	for _, fieldIdx := range colFieldsIndex {
		pt.ColFields.Field = append(pt.ColFields.Field, &xlsxField{
			X: fieldIdx,
		})
	}

	//in order to create pivot in case there is many Columns and Many Datas
	if len(opt.Data) > 1 {
		pt.ColFields.Field = append(pt.ColFields.Field, &xlsxField{
			X: -2,
		})
	}

	// count col fields
	pt.ColFields.Count = len(pt.ColFields.Field)
	return err
}

// addPivotFields create pivot fields based on the column order of the first
// row in the data region by given pivot table definition and option.
func (f *File) addPivotFields(pt *xlsxPivotTableDefinition, opt *PivotTableOption) error {
	order, err := f.getPivotFieldsOrder(opt)
	if err != nil {
		return err
	}
	x := 0
	for _, name := range order {
		if inPivotTableField(opt.Rows, name) != -1 {
			rowOptions, ok := f.getPivotTableFieldOptions(name, opt.Rows)
			var items []*xlsxItem
			if !ok || !rowOptions.DefaultSubtotal {
				items = append(items, &xlsxItem{X: &x})
			} else {
				items = append(items, &xlsxItem{T: "default"})
			}

			pt.PivotFields.PivotField = append(pt.PivotFields.PivotField, &xlsxPivotField{
				Name:            f.getPivotTableFieldName(name, opt.Rows),
				Axis:            "axisRow",
				Compact:         &rowOptions.Compact,
				Outline:         &rowOptions.Outline,
				DefaultSubtotal: &rowOptions.DefaultSubtotal,
				Items: &xlsxItems{
					Count: len(items),
					Item:  items,
				},
			})
			continue
		}
		if inPivotTableField(opt.Filter, name) != -1 {
			pt.PivotFields.PivotField = append(pt.PivotFields.PivotField, &xlsxPivotField{
				Axis: "axisPage",
				Name: f.getPivotTableFieldName(name, opt.Columns),
				Items: &xlsxItems{
					Count: 1,
					Item: []*xlsxItem{
						{T: "default"},
					},
				},
			})
			continue
		}
		if inPivotTableField(opt.Columns, name) != -1 {
			columnOptions, ok := f.getPivotTableFieldOptions(name, opt.Columns)
			var items []*xlsxItem
			if !ok || !columnOptions.DefaultSubtotal {
				items = append(items, &xlsxItem{X: &x})
			} else {
				items = append(items, &xlsxItem{T: "default"})
			}
			pt.PivotFields.PivotField = append(pt.PivotFields.PivotField, &xlsxPivotField{
				Name:            f.getPivotTableFieldName(name, opt.Columns),
				Axis:            "axisCol",
				Compact:         &columnOptions.Compact,
				Outline:         &columnOptions.Outline,
				DefaultSubtotal: &columnOptions.DefaultSubtotal,
				Items: &xlsxItems{
					Count: len(items),
					Item:  items,
				},
			})
			continue
		}
		if inPivotTableField(opt.Data, name) != -1 {
			pt.PivotFields.PivotField = append(pt.PivotFields.PivotField, &xlsxPivotField{
				DataField: true,
			})
			continue
		}
		pt.PivotFields.PivotField = append(pt.PivotFields.PivotField, &xlsxPivotField{})
	}
	return err
}

// countPivotTables provides a function to get drawing files count storage in
// the folder xl/pivotTables.
func (f *File) countPivotTables() int {
	count := 0
	f.Pkg.Range(func(k, v interface{}) bool {
		if strings.Contains(k.(string), "xl/pivotTables/pivotTable") {
			count++
		}
		return true
	})
	return count
}

// countPivotCache provides a function to get drawing files count storage in
// the folder xl/pivotCache.
func (f *File) countPivotCache() int {
	count := 0
	f.Pkg.Range(func(k, v interface{}) bool {
		if strings.Contains(k.(string), "xl/pivotCache/pivotCacheDefinition") {
			count++
		}
		return true
	})
	return count
}

// getPivotFieldsIndex convert the column of the first row in the data region
// to a sequential index by given fields and pivot option.
func (f *File) getPivotFieldsIndex(fields []PivotTableField, opt *PivotTableOption) ([]int, error) {
	pivotFieldsIndex := []int{}
	orders, err := f.getPivotFieldsOrder(opt)
	if err != nil {
		return pivotFieldsIndex, err
	}
	for _, field := range fields {
		if pos := inStrSlice(orders, field.Data); pos != -1 {
			pivotFieldsIndex = append(pivotFieldsIndex, pos)
		}
	}
	return pivotFieldsIndex, nil
}

// getPivotTableFieldsSubtotal prepare fields subtotal by given pivot table fields.
func (f *File) getPivotTableFieldsSubtotal(fields []PivotTableField) []string {
	field := make([]string, len(fields))
	enums := []string{"average", "count", "countNums", "max", "min", "product", "stdDev", "stdDevp", "sum", "var", "varp"}
	inEnums := func(enums []string, val string) string {
		for _, enum := range enums {
			if strings.EqualFold(enum, val) {
				return enum
			}
		}
		return "sum"
	}
	for idx, fld := range fields {
		field[idx] = inEnums(enums, fld.Subtotal)
	}
	return field
}

// getPivotTableFieldsName prepare fields name list by given pivot table
// fields.
func (f *File) getPivotTableFieldsName(fields []PivotTableField) []string {
	field := make([]string, len(fields))
	for idx, fld := range fields {
		if len(fld.Name) > MaxFieldLength {
			field[idx] = fld.Name[:MaxFieldLength]
			continue
		}
		field[idx] = fld.Name
	}
	return field
}

// getPivotTableFieldName prepare field name by given pivot table fields.
func (f *File) getPivotTableFieldName(name string, fields []PivotTableField) string {
	fieldsName := f.getPivotTableFieldsName(fields)
	for idx, field := range fields {
		if field.Data == name {
			return fieldsName[idx]
		}
	}
	return ""
}

// getPivotTableFieldOptions return options for specific field by given field name.
func (f *File) getPivotTableFieldOptions(name string, fields []PivotTableField) (options PivotTableField, ok bool) {
	for _, field := range fields {
		if field.Data == name {
			options, ok = field, true
			return
		}
	}
	return
}

// addWorkbookPivotCache add the association ID of the pivot cache in workbook.xml.
func (f *File) addWorkbookPivotCache(RID int) int {
	wb := f.workbookReader()
	if wb.PivotCaches == nil {
		wb.PivotCaches = &xlsxPivotCaches{}
	}
	cacheID := 1
	for _, pivotCache := range wb.PivotCaches.PivotCache {
		if pivotCache.CacheID > cacheID {
			cacheID = pivotCache.CacheID
		}
	}
	cacheID++
	wb.PivotCaches.PivotCache = append(wb.PivotCaches.PivotCache, xlsxPivotCache{
		CacheID: cacheID,
		RID:     fmt.Sprintf("rId%d", RID),
	})
	return cacheID
}
