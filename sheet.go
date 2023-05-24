package xlsx

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/mohae/deepcopy"
)

// NewSheet provides the function to create a new sheet by given a worksheet
// name and returns the index of the sheets in the workbook
// (spreadsheet) after it appended. Note that the worksheet names are not
// case sensitive, when creating a new spreadsheet file, the default
// worksheet named `Sheet1` will be created.
func (f *File) NewSheet(name string) int {
	// Check if the worksheet already exists
	index := f.GetSheetIndex(name)
	if index != -1 {
		return index
	}
	f.DeleteSheet(name)
	f.SheetCount++
	wb := f.workbookReader()
	sheetID := 0
	for _, v := range wb.Sheets.Sheet {
		if v.SheetID > sheetID {
			sheetID = v.SheetID
		}
	}
	sheetID++
	// Update docProps/app.xml
	f.setAppXML()
	// Update [Content_Types].xml
	f.setContentTypes("/xl/worksheets/sheet"+strconv.Itoa(sheetID)+".xml", ContentTypeSpreadSheetMLWorksheet)
	// Create new sheet /xl/worksheets/sheet%d.xml
	f.setSheet(sheetID, name)
	// Update workbook.xml.rels
	rID := f.addRels(f.getWorkbookRelsPath(), SourceRelationshipWorkSheet, fmt.Sprintf("/xl/worksheets/sheet%d.xml", sheetID), "")
	// Update workbook.xml
	f.setWorkbook(name, sheetID, rID)
	return f.GetSheetIndex(name)
}

// contentTypesReader provides a function to get the pointer to the
// [Content_Types].xml structure after deserialization.
func (f *File) contentTypesReader() *xlsxTypes {
	var err error

	if f.ContentTypes == nil {
		f.ContentTypes = new(xlsxTypes)
		f.ContentTypes.Lock()
		defer f.ContentTypes.Unlock()
		if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML("[Content_Types].xml")))).
			Decode(f.ContentTypes); err != nil && err != io.EOF {
			log.Printf("xml decode error: %s", err)
		}
	}
	return f.ContentTypes
}

// contentTypesWriter provides a function to save [Content_Types].xml after
// serialize structure.
func (f *File) contentTypesWriter() {
	if f.ContentTypes != nil {
		output, _ := xml.Marshal(f.ContentTypes)
		f.saveFileList("[Content_Types].xml", output)
	}
}

// getWorkbookPath provides a function to get the path of the workbook.xml in
// the spreadsheet.
func (f *File) getWorkbookPath() (path string) {
	if rels := f.relsReader("_rels/.rels"); rels != nil {
		rels.Lock()
		defer rels.Unlock()
		for _, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipOfficeDocument {
				path = strings.TrimPrefix(rel.Target, "/")
				return
			}
		}
	}
	return
}

// getWorkbookRelsPath provides a function to get the path of the workbook.xml.rels
// in the spreadsheet.
func (f *File) getWorkbookRelsPath() (path string) {
	wbPath := f.getWorkbookPath()
	wbDir := filepath.Dir(wbPath)
	if wbDir == "." {
		path = "_rels/" + filepath.Base(wbPath) + ".rels"
		return
	}
	path = strings.TrimPrefix(filepath.Dir(wbPath)+"/_rels/"+filepath.Base(wbPath)+".rels", "/")
	return
}

// getWorksheetPath construct a target XML as xl/worksheets/sheet%d by split
// path, compatible with different types of relative paths in
// workbook.xml.rels, for example: worksheets/sheet%d.xml
// and /xl/worksheets/sheet%d.xml
func (f *File) getWorksheetPath(relTarget string) (path string) {
	path = filepath.ToSlash(strings.TrimPrefix(
		strings.Replace(filepath.Clean(fmt.Sprintf("%s/%s", filepath.Dir(f.getWorkbookPath()), relTarget)), "\\", "/", -1), "/"))
	if strings.HasPrefix(relTarget, "/") {
		path = filepath.ToSlash(strings.TrimPrefix(strings.Replace(filepath.Clean(relTarget), "\\", "/", -1), "/"))
	}
	return path
}

// workbookReader provides a function to get the pointer to the workbook.xml
// structure after deserialization.
func (f *File) workbookReader() *xlsxWorkbook {
	var err error
	if f.WorkBook == nil {
		wbPath := f.getWorkbookPath()
		f.WorkBook = new(xlsxWorkbook)
		if _, ok := f.xmlAttr[wbPath]; !ok {
			d := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(wbPath))))
			f.xmlAttr[wbPath] = append(f.xmlAttr[wbPath], getRootElement(d)...)
			f.addNameSpaces(wbPath, SourceRelationship)
		}
		if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(wbPath)))).
			Decode(f.WorkBook); err != nil && err != io.EOF {
			log.Printf("xml decode error: %s", err)
		}
	}
	return f.WorkBook
}

// workBookWriter provides a function to save workbook.xml after serialize
// structure.
func (f *File) workBookWriter() {
	if f.WorkBook != nil {
		output, _ := xml.Marshal(f.WorkBook)
		f.saveFileList(f.getWorkbookPath(), replaceRelationshipsBytes(f.replaceNameSpaceBytes(f.getWorkbookPath(), output)))
	}
}

// mergeExpandedCols merge expanded columns.
func (f *File) mergeExpandedCols(ws *xlsxWorksheet) {
	sort.Slice(ws.Cols.Col, func(i, j int) bool {
		return ws.Cols.Col[i].Min < ws.Cols.Col[j].Min
	})
	columns := []xlsxCol{}
	for i, n := 0, len(ws.Cols.Col); i < n; {
		left := i
		for i++; i < n && reflect.DeepEqual(
			xlsxCol{
				BestFit:      ws.Cols.Col[i-1].BestFit,
				Collapsed:    ws.Cols.Col[i-1].Collapsed,
				CustomWidth:  ws.Cols.Col[i-1].CustomWidth,
				Hidden:       ws.Cols.Col[i-1].Hidden,
				Max:          ws.Cols.Col[i-1].Max + 1,
				Min:          ws.Cols.Col[i-1].Min + 1,
				OutlineLevel: ws.Cols.Col[i-1].OutlineLevel,
				Phonetic:     ws.Cols.Col[i-1].Phonetic,
				Style:        ws.Cols.Col[i-1].Style,
				Width:        ws.Cols.Col[i-1].Width,
			}, ws.Cols.Col[i]); i++ {
		}
		column := deepcopy.Copy(ws.Cols.Col[left]).(xlsxCol)
		if left < i-1 {
			column.Max = ws.Cols.Col[i-1].Min
		}
		columns = append(columns, column)
	}
	ws.Cols.Col = columns
}

// workSheetWriter provides a function to save xl/worksheets/sheet%d.xml after
// serialize structure.
func (f *File) workSheetWriter() {
	var arr []byte
	buffer := bytes.NewBuffer(arr)
	encoder := xml.NewEncoder(buffer)
	f.Sheet.Range(func(p, ws interface{}) bool {
		if ws != nil {
			sheet := ws.(*xlsxWorksheet)
			if sheet.MergeCells != nil && len(sheet.MergeCells.Cells) > 0 {
				_ = f.mergeOverlapCells(sheet)
			}
			if sheet.Cols != nil && len(sheet.Cols.Col) > 0 {
				f.mergeExpandedCols(sheet)
			}
			for k, v := range sheet.SheetData.Row {
				sheet.SheetData.Row[k].C = trimCell(v.C)
			}
			if sheet.SheetPr != nil || sheet.Drawing != nil || sheet.Hyperlinks != nil || sheet.Picture != nil || sheet.TableParts != nil {
				f.addNameSpaces(p.(string), SourceRelationship)
			}
			// reusing buffer
			_ = encoder.Encode(sheet)
			f.saveFileList(p.(string), replaceRelationshipsBytes(f.replaceNameSpaceBytes(p.(string), buffer.Bytes())))
			ok := f.checked[p.(string)]
			if ok {
				f.Sheet.Delete(p.(string))
				f.checked[p.(string)] = false
			}
			buffer.Reset()
		}
		return true
	})
}

// trimCell provides a function to trim blank cells which created by fillColumns.
func trimCell(column []xlsxC) []xlsxC {
	rowFull := true
	for i := range column {
		rowFull = column[i].hasValue() && rowFull
	}
	if rowFull {
		return column
	}
	col := make([]xlsxC, len(column))
	i := 0
	for _, c := range column {
		if c.hasValue() {
			col[i] = c
			i++
		}
	}
	return col[:i]
}

// setContentTypes provides a function to read and update property of contents
// type of the spreadsheet.
func (f *File) setContentTypes(partName, contentType string) {
	content := f.contentTypesReader()
	content.Lock()
	defer content.Unlock()
	content.Overrides = append(content.Overrides, xlsxOverride{
		PartName:    partName,
		ContentType: contentType,
	})
}

// setSheet provides a function to update sheet property by given index.
func (f *File) setSheet(index int, name string) {
	ws := xlsxWorksheet{
		Dimension: &xlsxDimension{Ref: "A1"},
		SheetViews: &xlsxSheetViews{
			SheetView: []xlsxSheetView{{WorkbookViewID: 0}},
		},
	}
	path := "xl/worksheets/sheet" + strconv.Itoa(index) + ".xml"
	f.sheetMap[trimSheetName(name)] = path
	f.Sheet.Store(path, &ws)
	f.xmlAttr[path] = []xml.Attr{NameSpaceSpreadSheet}
}

// setWorkbook update workbook property of the spreadsheet. Maximum 31
// characters are allowed in sheet title.
func (f *File) setWorkbook(name string, sheetID, rid int) {
	content := f.workbookReader()
	content.Sheets.Sheet = append(content.Sheets.Sheet, xlsxSheet{
		Name:    trimSheetName(name),
		SheetID: sheetID,
		ID:      "rId" + strconv.Itoa(rid),
	})
}

// relsWriter provides a function to save relationships after
// serialize structure.
func (f *File) relsWriter() {
	f.Relationships.Range(func(path, rel interface{}) bool {
		if rel != nil {
			output, _ := xml.Marshal(rel.(*xlsxRelationships))
			if strings.HasPrefix(path.(string), "xl/worksheets/sheet/rels/sheet") {
				output = f.replaceNameSpaceBytes(path.(string), output)
			}
			f.saveFileList(path.(string), replaceRelationshipsBytes(output))
		}
		return true
	})
}

// setAppXML update docProps/app.xml file of XML.
func (f *File) setAppXML() {
	f.saveFileList("docProps/app.xml", []byte(templateDocpropsApp))
}

// replaceRelationshipsBytes; Some tools that read spreadsheet files have very
// strict requirements about the structure of the input XML. This function is
// a horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsBytes(content []byte) []byte {
	oldXmlns := []byte(`xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships`)
	newXmlns := []byte("r")
	return bytesReplace(content, oldXmlns, newXmlns, -1)
}

// SetActiveSheet provides a function to set the default active sheet of the
// workbook by a given index. Note that the active index is different from the
// ID returned by function GetSheetMap(). It should be greater or equal to 0
// and less than the total worksheet numbers.
func (f *File) SetActiveSheet(index int) {
	if index < 0 {
		index = 0
	}
	wb := f.workbookReader()
	for activeTab := range wb.Sheets.Sheet {
		if activeTab == index {
			if wb.BookViews == nil {
				wb.BookViews = &xlsxBookViews{}
			}
			if len(wb.BookViews.WorkBookView) > 0 {
				wb.BookViews.WorkBookView[0].ActiveTab = activeTab
			} else {
				wb.BookViews.WorkBookView = append(wb.BookViews.WorkBookView, xlsxWorkBookView{
					ActiveTab: activeTab,
				})
			}
		}
	}
	for idx, name := range f.GetSheetList() {
		ws, err := f.workSheetReader(name)
		if err != nil {
			// Chartsheet or dialogsheet
			return
		}
		if ws.SheetViews == nil {
			ws.SheetViews = &xlsxSheetViews{
				SheetView: []xlsxSheetView{{WorkbookViewID: 0}},
			}
		}
		if len(ws.SheetViews.SheetView) > 0 {
			ws.SheetViews.SheetView[0].TabSelected = false
		}
		if index == idx {
			if len(ws.SheetViews.SheetView) > 0 {
				ws.SheetViews.SheetView[0].TabSelected = true
			} else {
				ws.SheetViews.SheetView = append(ws.SheetViews.SheetView, xlsxSheetView{
					TabSelected: true,
				})
			}
		}
	}
}

// GetActiveSheetIndex provides a function to get active sheet index of the
// spreadsheet. If not found the active sheet will be return integer 0.
func (f *File) GetActiveSheetIndex() (index int) {
	var sheetID = f.getActiveSheetID()
	wb := f.workbookReader()
	if wb != nil {
		for idx, sheet := range wb.Sheets.Sheet {
			if sheet.SheetID == sheetID {
				index = idx
				return
			}
		}
	}
	return
}

// getActiveSheetID provides a function to get active sheet ID of the
// spreadsheet. If not found the active sheet will be return integer 0.
func (f *File) getActiveSheetID() int {
	wb := f.workbookReader()
	if wb != nil {
		if wb.BookViews != nil && len(wb.BookViews.WorkBookView) > 0 {
			activeTab := wb.BookViews.WorkBookView[0].ActiveTab
			if len(wb.Sheets.Sheet) > activeTab && wb.Sheets.Sheet[activeTab].SheetID != 0 {
				return wb.Sheets.Sheet[activeTab].SheetID
			}
		}
		if len(wb.Sheets.Sheet) >= 1 {
			return wb.Sheets.Sheet[0].SheetID
		}
	}
	return 0
}

// SetSheetName provides a function to set the worksheet name by given old and
// new worksheet names. Maximum 31 characters are allowed in sheet title and
// this function only changes the name of the sheet and will not update the
// sheet name in the formula or reference associated with the cell. So there
// may be problem formula error or reference missing.
func (f *File) SetSheetName(oldName, newName string) {
	oldName = trimSheetName(oldName)
	newName = trimSheetName(newName)
	if newName == oldName {
		return
	}
	content := f.workbookReader()
	for k, v := range content.Sheets.Sheet {
		if v.Name == oldName {
			content.Sheets.Sheet[k].Name = newName
			f.sheetMap[newName] = f.sheetMap[oldName]
			delete(f.sheetMap, oldName)
		}
	}
}

// GetSheetName provides a function to get the sheet name of the workbook by
// the given sheet index. If the given sheet index is invalid, it will return
// an empty string.
func (f *File) GetSheetName(index int) (name string) {
	for idx, sheet := range f.GetSheetList() {
		if idx == index {
			name = sheet
			return
		}
	}
	return
}

// getSheetID provides a function to get worksheet ID of the spreadsheet by
// given sheet name. If given worksheet name is invalid, will return an
// integer type value -1.
func (f *File) getSheetID(name string) int {
	for sheetID, sheet := range f.GetSheetMap() {
		if sheet == trimSheetName(name) {
			return sheetID
		}
	}
	return -1
}

// GetSheetIndex provides a function to get a sheet index of the workbook by
// the given sheet name, the sheet names are not case sensitive. If the given
// sheet name is invalid or sheet doesn't exist, it will return an integer
// type value -1.
func (f *File) GetSheetIndex(name string) int {
	for index, sheet := range f.GetSheetList() {
		if strings.EqualFold(sheet, trimSheetName(name)) {
			return index
		}
	}
	return -1
}

// GetSheetMap provides a function to get worksheets, chart sheets, dialog
// sheets ID and name map of the workbook. For example:
//
//    f, err := xlsx.OpenFile("Book1.xlsx")
//    if err != nil {
//        return
//    }
//    defer func() {
//        if err := f.Close(); err != nil {
//            fmt.Println(err)
//        }
//    }()
//    for index, name := range f.GetSheetMap() {
//        fmt.Println(index, name)
//    }
//
func (f *File) GetSheetMap() map[int]string {
	wb := f.workbookReader()
	sheetMap := map[int]string{}
	if wb != nil {
		for _, sheet := range wb.Sheets.Sheet {
			sheetMap[sheet.SheetID] = sheet.Name
		}
	}
	return sheetMap
}

// GetSheetList provides a function to get worksheets, chart sheets, and
// dialog sheets name list of the workbook.
func (f *File) GetSheetList() (list []string) {
	wb := f.workbookReader()
	if wb != nil {
		for _, sheet := range wb.Sheets.Sheet {
			list = append(list, sheet.Name)
		}
	}
	return
}

// getSheetMap provides a function to get worksheet name and XML file path map
// of the spreadsheet.
func (f *File) getSheetMap() map[string]string {
	maps := map[string]string{}
	for _, v := range f.workbookReader().Sheets.Sheet {
		for _, rel := range f.relsReader(f.getWorkbookRelsPath()).Relationships {
			if rel.ID == v.ID {
				path := f.getWorksheetPath(rel.Target)
				if _, ok := f.Pkg.Load(path); ok {
					maps[v.Name] = path
				}
				if _, ok := f.tempFiles.Load(path); ok {
					maps[v.Name] = path
				}
			}
		}
	}
	return maps
}

// SetSheetBackground provides a function to set background picture by given
// worksheet name and file path.
func (f *File) SetSheetBackground(sheet, picture string) error {
	var err error
	// Check picture exists first.
	if _, err = os.Stat(picture); os.IsNotExist(err) {
		return err
	}
	ext, ok := supportImageTypes[path.Ext(picture)]
	if !ok {
		return ErrImgExt
	}
	file, _ := ioutil.ReadFile(filepath.Clean(picture))
	name := f.addMedia(file, ext)
	sheetRels := "xl/worksheets/_rels/" + strings.TrimPrefix(f.sheetMap[trimSheetName(sheet)], "xl/worksheets/") + ".rels"
	rID := f.addRels(sheetRels, SourceRelationshipImage, strings.Replace(name, "xl", "..", 1), "")
	f.addSheetPicture(sheet, rID)
	f.addSheetNameSpace(sheet, SourceRelationship)
	f.setContentTypePartImageExtensions()
	return err
}

// DeleteSheet provides a function to delete worksheet in a workbook by given
// worksheet name, the sheet names are not case sensitive.the sheet names are
// not case sensitive. Use this method with caution, which will affect
// changes in references such as formulas, charts, and so on. If there is any
// referenced value of the deleted worksheet, it will cause a file error when
// you open it. This function will be invalid when only the one worksheet is
// left.
func (f *File) DeleteSheet(name string) {
	if f.SheetCount == 1 || f.GetSheetIndex(name) == -1 {
		return
	}
	sheetName := trimSheetName(name)
	wb := f.workbookReader()
	wbRels := f.relsReader(f.getWorkbookRelsPath())
	activeSheetName := f.GetSheetName(f.GetActiveSheetIndex())
	deleteLocalSheetID := f.GetSheetIndex(name)
	// Delete and adjust defined names
	if wb.DefinedNames != nil {
		for idx := 0; idx < len(wb.DefinedNames.DefinedName); idx++ {
			dn := wb.DefinedNames.DefinedName[idx]
			if dn.LocalSheetID != nil {
				localSheetID := *dn.LocalSheetID
				if localSheetID == deleteLocalSheetID {
					wb.DefinedNames.DefinedName = append(wb.DefinedNames.DefinedName[:idx], wb.DefinedNames.DefinedName[idx+1:]...)
					idx--
				} else if localSheetID > deleteLocalSheetID {
					wb.DefinedNames.DefinedName[idx].LocalSheetID = intPtr(*dn.LocalSheetID - 1)
				}
			}
		}
	}
	for idx, sheet := range wb.Sheets.Sheet {
		if strings.EqualFold(sheet.Name, sheetName) {
			wb.Sheets.Sheet = append(wb.Sheets.Sheet[:idx], wb.Sheets.Sheet[idx+1:]...)
			var sheetXML, rels string
			if wbRels != nil {
				for _, rel := range wbRels.Relationships {
					if rel.ID == sheet.ID {
						sheetXML = f.getWorksheetPath(rel.Target)
						rels = "xl/worksheets/_rels/" + strings.TrimPrefix(f.sheetMap[sheetName], "xl/worksheets/") + ".rels"
					}
				}
			}
			target := f.deleteSheetFromWorkbookRels(sheet.ID)
			f.deleteSheetFromContentTypes(target)
			f.deleteCalcChain(sheet.SheetID, "")
			delete(f.sheetMap, sheet.Name)
			f.Pkg.Delete(sheetXML)
			f.Pkg.Delete(rels)
			f.Relationships.Delete(rels)
			f.Sheet.Delete(sheetXML)
			delete(f.xmlAttr, sheetXML)
			f.SheetCount--
		}
	}
	f.SetActiveSheet(f.GetSheetIndex(activeSheetName))
}

// deleteSheetFromWorkbookRels provides a function to remove worksheet
// relationships by given relationships ID in the file workbook.xml.rels.
func (f *File) deleteSheetFromWorkbookRels(rID string) string {
	content := f.relsReader(f.getWorkbookRelsPath())
	content.Lock()
	defer content.Unlock()
	for k, v := range content.Relationships {
		if v.ID == rID {
			content.Relationships = append(content.Relationships[:k], content.Relationships[k+1:]...)
			return v.Target
		}
	}
	return ""
}

// deleteSheetFromContentTypes provides a function to remove worksheet
// relationships by given target name in the file [Content_Types].xml.
func (f *File) deleteSheetFromContentTypes(target string) {
	if !strings.HasPrefix(target, "/") {
		target = "/xl/" + target
	}
	content := f.contentTypesReader()
	content.Lock()
	defer content.Unlock()
	for k, v := range content.Overrides {
		if v.PartName == target {
			content.Overrides = append(content.Overrides[:k], content.Overrides[k+1:]...)
		}
	}
}

// CopySheet provides a function to duplicate a worksheet by gave source and
// target worksheet index. Note that currently doesn't support duplicate
// workbooks that contain tables, charts or pictures. For Example:
//
//    // Sheet1 already exists...
//    index := f.NewSheet("Sheet2")
//    err := f.CopySheet(1, index)
//    return err
//
func (f *File) CopySheet(from, to int) error {
	if from < 0 || to < 0 || from == to || f.GetSheetName(from) == "" || f.GetSheetName(to) == "" {
		return ErrSheetIdx
	}
	return f.copySheet(from, to)
}

// copySheet provides a function to duplicate a worksheet by gave source and
// target worksheet name.
func (f *File) copySheet(from, to int) error {
	fromSheet := f.GetSheetName(from)
	sheet, err := f.workSheetReader(fromSheet)
	if err != nil {
		return err
	}
	worksheet := deepcopy.Copy(sheet).(*xlsxWorksheet)
	toSheetID := strconv.Itoa(f.getSheetID(f.GetSheetName(to)))
	path := "xl/worksheets/sheet" + toSheetID + ".xml"
	if len(worksheet.SheetViews.SheetView) > 0 {
		worksheet.SheetViews.SheetView[0].TabSelected = false
	}
	worksheet.Drawing = nil
	worksheet.TableParts = nil
	worksheet.PageSetUp = nil
	f.Sheet.Store(path, worksheet)
	toRels := "xl/worksheets/_rels/sheet" + toSheetID + ".xml.rels"
	fromRels := "xl/worksheets/_rels/sheet" + strconv.Itoa(f.getSheetID(fromSheet)) + ".xml.rels"
	if rels, ok := f.Pkg.Load(fromRels); ok && rels != nil {
		f.Pkg.Store(toRels, rels.([]byte))
	}
	fromSheetXMLPath := f.sheetMap[trimSheetName(fromSheet)]
	fromSheetAttr := f.xmlAttr[fromSheetXMLPath]
	f.xmlAttr[path] = fromSheetAttr
	return err
}

// SetSheetVisible provides a function to set worksheet visible by given worksheet
// name. A workbook must contain at least one visible worksheet. If the given
// worksheet has been activated, this setting will be invalidated. Sheet state
// values as defined by https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetstatevalues
//
//    visible
//    hidden
//    veryHidden
//
// For example, hide Sheet1:
//
//    err := f.SetSheetVisible("Sheet1", false)
//
func (f *File) SetSheetVisible(name string, visible bool) error {
	name = trimSheetName(name)
	content := f.workbookReader()
	if visible {
		for k, v := range content.Sheets.Sheet {
			if v.Name == name {
				content.Sheets.Sheet[k].State = ""
			}
		}
		return nil
	}
	count := 0
	for _, v := range content.Sheets.Sheet {
		if v.State != "hidden" {
			count++
		}
	}
	for k, v := range content.Sheets.Sheet {
		ws, err := f.workSheetReader(v.Name)
		if err != nil {
			return err
		}
		tabSelected := false
		if len(ws.SheetViews.SheetView) > 0 {
			tabSelected = ws.SheetViews.SheetView[0].TabSelected
		}
		if v.Name == name && count > 1 && !tabSelected {
			content.Sheets.Sheet[k].State = "hidden"
		}
	}
	return nil
}

// parseFormatPanesSet provides a function to parse the panes settings.
func parseFormatPanesSet(formatSet string) (*formatPanes, error) {
	format := formatPanes{}
	err := json.Unmarshal([]byte(formatSet), &format)
	return &format, err
}

// SetPanes provides a function to create and remove freeze panes and split panes
// by given worksheet name and panes format set.
//
// activePane defines the pane that is active. The possible values for this
// attribute are defined in the following table:
//
//     Enumeration Value              | Description
//    --------------------------------+-------------------------------------------------------------
//     bottomLeft (Bottom Left Pane)  | Bottom left pane, when both vertical and horizontal
//                                    | splits are applied.
//                                    |
//                                    | This value is also used when only a horizontal split has
//                                    | been applied, dividing the pane into upper and lower
//                                    | regions. In that case, this value specifies the bottom
//                                    | pane.
//                                    |
//    bottomRight (Bottom Right Pane) | Bottom right pane, when both vertical and horizontal
//                                    | splits are applied.
//                                    |
//    topLeft (Top Left Pane)         | Top left pane, when both vertical and horizontal splits
//                                    | are applied.
//                                    |
//                                    | This value is also used when only a horizontal split has
//                                    | been applied, dividing the pane into upper and lower
//                                    | regions. In that case, this value specifies the top pane.
//                                    |
//                                    | This value is also used when only a vertical split has
//                                    | been applied, dividing the pane into right and left
//                                    | regions. In that case, this value specifies the left pane
//                                    |
//    topRight (Top Right Pane)       | Top right pane, when both vertical and horizontal
//                                    | splits are applied.
//                                    |
//                                    | This value is also used when only a vertical split has
//                                    | been applied, dividing the pane into right and left
//                                    | regions. In that case, this value specifies the right
//                                    | pane.
//
// Pane state type is restricted to the values supported currently listed in the following table:
//
//     Enumeration Value              | Description
//    --------------------------------+-------------------------------------------------------------
//     frozen (Frozen)                | Panes are frozen, but were not split being frozen. In
//                                    | this state, when the panes are unfrozen again, a single
//                                    | pane results, with no split.
//                                    |
//                                    | In this state, the split bars are not adjustable.
//                                    |
//     split (Split)                  | Panes are split, but not frozen. In this state, the split
//                                    | bars are adjustable by the user.
//
// x_split (Horizontal Split Position): Horizontal position of the split, in
// 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value
// indicates the number of columns visible in the top pane.
//
// y_split (Vertical Split Position): Vertical position of the split, in 1/20th
// of a point; 0 (zero) if none. If the pane is frozen, this value indicates the
// number of rows visible in the left pane. The possible values for this
// attribute are defined by the W3C XML Schema double datatype.
//
// top_left_cell: Location of the top left visible cell in the bottom right pane
// (when in Left-To-Right mode).
//
// sqref (Sequence of References): Range of the selection. Can be non-contiguous
// set of ranges.
//
// An example of how to freeze column A in the Sheet1 and set the active cell on
// Sheet1!K16:
//
//    f.SetPanes("Sheet1", `{"freeze":true,"split":false,"x_split":1,"y_split":0,"top_left_cell":"B1","active_pane":"topRight","panes":[{"sqref":"K16","active_cell":"K16","pane":"topRight"}]}`)
//
// An example of how to freeze rows 1 to 9 in the Sheet1 and set the active cell
// ranges on Sheet1!A11:XFD11:
//
//    f.SetPanes("Sheet1", `{"freeze":true,"split":false,"x_split":0,"y_split":9,"top_left_cell":"A34","active_pane":"bottomLeft","panes":[{"sqref":"A11:XFD11","active_cell":"A11","pane":"bottomLeft"}]}`)
//
// An example of how to create split panes in the Sheet1 and set the active cell
// on Sheet1!J60:
//
//    f.SetPanes("Sheet1", `{"freeze":false,"split":true,"x_split":3270,"y_split":1800,"top_left_cell":"N57","active_pane":"bottomLeft","panes":[{"sqref":"I36","active_cell":"I36"},{"sqref":"G33","active_cell":"G33","pane":"topRight"},{"sqref":"J60","active_cell":"J60","pane":"bottomLeft"},{"sqref":"O60","active_cell":"O60","pane":"bottomRight"}]}`)
//
// An example of how to unfreeze and remove all panes on Sheet1:
//
//    f.SetPanes("Sheet1", `{"freeze":false,"split":false}`)
//
func (f *File) SetPanes(sheet, panes string) error {
	fs, _ := parseFormatPanesSet(panes)
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	p := &xlsxPane{
		ActivePane:  fs.ActivePane,
		TopLeftCell: fs.TopLeftCell,
		XSplit:      float64(fs.XSplit),
		YSplit:      float64(fs.YSplit),
	}
	if fs.Freeze {
		p.State = "frozen"
	}
	ws.SheetViews.SheetView[len(ws.SheetViews.SheetView)-1].Pane = p
	if !(fs.Freeze) && !(fs.Split) {
		if len(ws.SheetViews.SheetView) > 0 {
			ws.SheetViews.SheetView[len(ws.SheetViews.SheetView)-1].Pane = nil
		}
	}
	s := []*xlsxSelection{}
	for _, p := range fs.Panes {
		s = append(s, &xlsxSelection{
			ActiveCell: p.ActiveCell,
			Pane:       p.Pane,
			SQRef:      p.SQRef,
		})
	}
	ws.SheetViews.SheetView[len(ws.SheetViews.SheetView)-1].Selection = s
	return err
}

// GetSheetVisible provides a function to get worksheet visible by given worksheet
// name. For example, get visible state of Sheet1:
//
//    f.GetSheetVisible("Sheet1")
//
func (f *File) GetSheetVisible(name string) bool {
	content := f.workbookReader()
	visible := false
	for k, v := range content.Sheets.Sheet {
		if v.Name == trimSheetName(name) {
			if content.Sheets.Sheet[k].State == "" || content.Sheets.Sheet[k].State == "visible" {
				visible = true
			}
		}
	}
	return visible
}

// SearchSheet provides a function to get coordinates by given worksheet name,
// cell value, and regular expression. The function doesn't support searching
// on the calculated result, formatted numbers and conditional lookup
// currently. If it is a merged cell, it will return the coordinates of the
// upper left corner of the merged area.
//
// An example of search the coordinates of the value of "100" on Sheet1:
//
//    result, err := f.SearchSheet("Sheet1", "100")
//
// An example of search the coordinates where the numerical value in the range
// of "0-9" of Sheet1 is described:
//
//    result, err := f.SearchSheet("Sheet1", "[0-9]", true)
//
func (f *File) SearchSheet(sheet, value string, reg ...bool) ([]string, error) {
	var (
		regSearch bool
		result    []string
	)
	for _, r := range reg {
		regSearch = r
	}
	name, ok := f.sheetMap[trimSheetName(sheet)]
	if !ok {
		return result, ErrSheetNotExist{sheet}
	}
	if ws, ok := f.Sheet.Load(name); ok && ws != nil {
		// flush data
		output, _ := xml.Marshal(ws.(*xlsxWorksheet))
		f.saveFileList(name, f.replaceNameSpaceBytes(name, output))
	}
	return f.searchSheet(name, value, regSearch)
}

// searchSheet provides a function to get coordinates by given worksheet name,
// cell value, and regular expression.
func (f *File) searchSheet(name, value string, regSearch bool) (result []string, err error) {
	var (
		cellName, inElement string
		cellCol, row        int
		d                   *xlsxSST
	)

	d = f.sharedStringsReader()
	decoder := f.xmlNewDecoder(bytes.NewReader(f.readBytes(name)))
	for {
		var token xml.Token
		token, err = decoder.Token()
		if err != nil || token == nil {
			if err == io.EOF {
				err = nil
			}
			break
		}
		switch xmlElement := token.(type) {
		case xml.StartElement:
			inElement = xmlElement.Name.Local
			if inElement == "row" {
				row, err = attrValToInt("r", xmlElement.Attr)
				if err != nil {
					return
				}
			}
			if inElement == "c" {
				colCell := xlsxC{}
				_ = decoder.DecodeElement(&colCell, &xmlElement)
				val, _ := colCell.getValueFrom(f, d, false)
				if regSearch {
					regex := regexp.MustCompile(value)
					if !regex.MatchString(val) {
						continue
					}
				} else {
					if val != value {
						continue
					}
				}
				cellCol, _, err = CellNameToCoordinates(colCell.R)
				if err != nil {
					return result, err
				}
				cellName, err = CoordinatesToCellName(cellCol, row)
				if err != nil {
					return result, err
				}
				result = append(result, cellName)
			}
		default:
		}
	}
	return
}

// attrValToInt provides a function to convert the local names to an integer
// by given XML attributes and specified names.
func attrValToInt(name string, attrs []xml.Attr) (val int, err error) {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			val, err = strconv.Atoi(attr.Value)
			if err != nil {
				return
			}
		}
	}
	return
}

// SetHeaderFooter provides a function to set headers and footers by given
// worksheet name and the control characters.
//
// Headers and footers are specified using the following settings fields:
//
//     Fields           | Description
//    ------------------+-----------------------------------------------------------
//     AlignWithMargins | Align header footer margins with page margins
//     DifferentFirst   | Different first-page header and footer indicator
//     DifferentOddEven | Different odd and even page headers and footers indicator
//     ScaleWithDoc     | Scale header and footer with document scaling
//     OddFooter        | Odd Page Footer
//     OddHeader        | Odd Header
//     EvenFooter       | Even Page Footer
//     EvenHeader       | Even Page Header
//     FirstFooter      | First Page Footer
//     FirstHeader      | First Page Header
//
// The following formatting codes can be used in 6 string type fields:
// OddHeader, OddFooter, EvenHeader, EvenFooter, FirstFooter, FirstHeader
//
//     Formatting Code        | Description
//    ------------------------+-------------------------------------------------------------------------
//     &&                     | The character "&"
//                            |
//     &font-size             | Size of the text font, where font-size is a decimal font size in points
//                            |
//     &"font name,font type" | A text font-name string, font name, and a text font-type string,
//                            | font type
//                            |
//     &"-,Regular"           | Regular text format. Toggles bold and italic modes to off
//                            |
//     &A                     | Current worksheet's tab name
//                            |
//     &B or &"-,Bold"        | Bold text format, from off to on, or vice versa. The default mode is off
//                            |
//     &D                     | Current date
//                            |
//     &C                     | Center section
//                            |
//     &E                     | Double-underline text format
//                            |
//     &F                     | Current workbook's file name
//                            |
//     &G                     | Drawing object as background
//                            |
//     &H                     | Shadow text format
//                            |
//     &I or &"-,Italic"      | Italic text format
//                            |
//     &K                     | Text font color
//                            |
//                            | An RGB Color is specified as RRGGBB
//                            |
//                            | A Theme Color is specified as TTSNNN where TT is the theme color Id,
//                            | S is either "+" or "-" of the tint/shade value, and NNN is the
//                            | tint/shade value
//                            |
//     &L                     | Left section
//                            |
//     &N                     | Total number of pages
//                            |
//     &O                     | Outline text format
//                            |
//     &P[[+|-]n]             | Without the optional suffix, the current page number in decimal
//                            |
//     &R                     | Right section
//                            |
//     &S                     | Strikethrough text format
//                            |
//     &T                     | Current time
//                            |
//     &U                     | Single-underline text format. If double-underline mode is on, the next
//                            | occurrence in a section specifier toggles double-underline mode to off;
//                            | otherwise, it toggles single-underline mode, from off to on, or vice
//                            | versa. The default mode is off
//                            |
//     &X                     | Superscript text format
//                            |
//     &Y                     | Subscript text format
//                            |
//     &Z                     | Current workbook's file path
//
// For example:
//
//    err := f.SetHeaderFooter("Sheet1", &xlsx.FormatHeaderFooter{
//        DifferentFirst:   true,
//        DifferentOddEven: true,
//        OddHeader:        "&R&P",
//        OddFooter:        "&C&F",
//        EvenHeader:       "&L&P",
//        EvenFooter:       "&L&D&R&T",
//        FirstHeader:      `&CCenter &"-,Bold"Bold&"-,Regular"HeaderU+000A&D`,
//    })
//
// This example shows:
//
// - The first page has its own header and footer
//
// - Odd and even-numbered pages have different headers and footers
//
// - Current page number in the right section of odd-page headers
//
// - Current workbook's file name in the center section of odd-page footers
//
// - Current page number in the left section of even-page headers
//
// - Current date in the left section and the current time in the right section
// of even-page footers
//
// - The text "Center Bold Header" on the first line of the center section of
// the first page, and the date on the second line of the center section of
// that same page
//
// - No footer on the first page
//
func (f *File) SetHeaderFooter(sheet string, settings *FormatHeaderFooter) error {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	if settings == nil {
		ws.HeaderFooter = nil
		return err
	}

	v := reflect.ValueOf(*settings)
	// Check 6 string type fields: OddHeader, OddFooter, EvenHeader, EvenFooter,
	// FirstFooter, FirstHeader
	for i := 4; i < v.NumField()-1; i++ {
		if len(utf16.Encode([]rune(v.Field(i).String()))) > MaxFieldLength {
			return newFieldLengthError(v.Type().Field(i).Name)
		}
	}
	ws.HeaderFooter = &xlsxHeaderFooter{
		AlignWithMargins: settings.AlignWithMargins,
		DifferentFirst:   settings.DifferentFirst,
		DifferentOddEven: settings.DifferentOddEven,
		ScaleWithDoc:     settings.ScaleWithDoc,
		OddHeader:        settings.OddHeader,
		OddFooter:        settings.OddFooter,
		EvenHeader:       settings.EvenHeader,
		EvenFooter:       settings.EvenFooter,
		FirstFooter:      settings.FirstFooter,
		FirstHeader:      settings.FirstHeader,
	}
	return err
}

// ProtectSheet provides a function to prevent other users from accidentally
// or deliberately changing, moving, or deleting data in a worksheet. For
// example, protect Sheet1 with protection settings:
//
//    err := f.ProtectSheet("Sheet1", &xlsx.FormatSheetProtection{
//        Password:      "password",
//        EditScenarios: false,
//    })
//
func (f *File) ProtectSheet(sheet string, settings *FormatSheetProtection) error {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	if settings == nil {
		settings = &FormatSheetProtection{
			EditObjects:       true,
			EditScenarios:     true,
			SelectLockedCells: true,
		}
	}
	ws.SheetProtection = &xlsxSheetProtection{
		AutoFilter:          settings.AutoFilter,
		DeleteColumns:       settings.DeleteColumns,
		DeleteRows:          settings.DeleteRows,
		FormatCells:         settings.FormatCells,
		FormatColumns:       settings.FormatColumns,
		FormatRows:          settings.FormatRows,
		InsertColumns:       settings.InsertColumns,
		InsertHyperlinks:    settings.InsertHyperlinks,
		InsertRows:          settings.InsertRows,
		Objects:             settings.EditObjects,
		PivotTables:         settings.PivotTables,
		Scenarios:           settings.EditScenarios,
		SelectLockedCells:   settings.SelectLockedCells,
		SelectUnlockedCells: settings.SelectUnlockedCells,
		Sheet:               true,
		Sort:                settings.Sort,
	}
	if settings.Password != "" {
		ws.SheetProtection.Password = genSheetPasswd(settings.Password)
	}
	return err
}

// UnprotectSheet provides a function to unprotect an Excel worksheet.
func (f *File) UnprotectSheet(sheet string) error {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	ws.SheetProtection = nil
	return err
}

// trimSheetName provides a function to trim invaild characters by given worksheet
// name.
func trimSheetName(name string) string {
	if strings.ContainsAny(name, ":\\/?*[]") || utf8.RuneCountInString(name) > 31 {
		r := make([]rune, 0, 31)
		for _, v := range name {
			switch v {
			case 58, 92, 47, 63, 42, 91, 93: // replace :\/?*[]
				continue
			default:
				r = append(r, v)
			}
			if len(r) == 31 {
				break
			}
		}
		name = string(r)
	}
	return name
}

// PageLayoutOption is an option of a page layout of a worksheet. See
// SetPageLayout().
type PageLayoutOption interface {
	setPageLayout(layout *xlsxPageSetUp)
}

// PageLayoutOptionPtr is a writable PageLayoutOption. See GetPageLayout().
type PageLayoutOptionPtr interface {
	PageLayoutOption
	getPageLayout(layout *xlsxPageSetUp)
}

type (
	// BlackAndWhite specified print black and white.
	BlackAndWhite bool
	// FirstPageNumber specified the first printed page number. If no value is
	// specified, then 'automatic' is assumed.
	FirstPageNumber uint
	// PageLayoutOrientation defines the orientation of page layout for a
	// worksheet.
	PageLayoutOrientation string
	// PageLayoutPaperSize defines the paper size of the worksheet.
	PageLayoutPaperSize int
	// FitToHeight specified the number of vertical pages to fit on.
	FitToHeight int
	// FitToWidth specified the number of horizontal pages to fit on.
	FitToWidth int
	// PageLayoutScale defines the print scaling. This attribute is restricted
	// to values ranging from 10 (10%) to 400 (400%). This setting is
	// overridden when fitToWidth and/or fitToHeight are in use.
	PageLayoutScale uint
)

const (
	// OrientationPortrait indicates page layout orientation id portrait.
	OrientationPortrait = "portrait"
	// OrientationLandscape indicates page layout orientation id landscape.
	OrientationLandscape = "landscape"
)

// setPageLayout provides a method to set the print black and white for the
// worksheet.
func (p BlackAndWhite) setPageLayout(ps *xlsxPageSetUp) {
	ps.BlackAndWhite = bool(p)
}

// getPageLayout provides a method to get the print black and white for the
// worksheet.
func (p *BlackAndWhite) getPageLayout(ps *xlsxPageSetUp) {
	if ps == nil {
		*p = false
		return
	}
	*p = BlackAndWhite(ps.BlackAndWhite)
}

// setPageLayout provides a method to set the first printed page number for
// the worksheet.
func (p FirstPageNumber) setPageLayout(ps *xlsxPageSetUp) {
	if 0 < int(p) {
		ps.FirstPageNumber = strconv.Itoa(int(p))
		ps.UseFirstPageNumber = true
	}
}

// getPageLayout provides a method to get the first printed page number for
// the worksheet.
func (p *FirstPageNumber) getPageLayout(ps *xlsxPageSetUp) {
	if ps != nil && ps.UseFirstPageNumber {
		if number, _ := strconv.Atoi(ps.FirstPageNumber); number != 0 {
			*p = FirstPageNumber(number)
			return
		}
	}
	*p = 1
}

// setPageLayout provides a method to set the orientation for the worksheet.
func (o PageLayoutOrientation) setPageLayout(ps *xlsxPageSetUp) {
	ps.Orientation = string(o)
}

// getPageLayout provides a method to get the orientation for the worksheet.
func (o *PageLayoutOrientation) getPageLayout(ps *xlsxPageSetUp) {
	// Excel default: portrait
	if ps == nil || ps.Orientation == "" {
		*o = OrientationPortrait
		return
	}
	*o = PageLayoutOrientation(ps.Orientation)
}

// setPageLayout provides a method to set the paper size for the worksheet.
func (p PageLayoutPaperSize) setPageLayout(ps *xlsxPageSetUp) {
	ps.PaperSize = int(p)
}

// getPageLayout provides a method to get the paper size for the worksheet.
func (p *PageLayoutPaperSize) getPageLayout(ps *xlsxPageSetUp) {
	// Excel default: 1
	if ps == nil || ps.PaperSize == 0 {
		*p = 1
		return
	}
	*p = PageLayoutPaperSize(ps.PaperSize)
}

// setPageLayout provides a method to set the fit to height for the worksheet.
func (p FitToHeight) setPageLayout(ps *xlsxPageSetUp) {
	if int(p) > 0 {
		ps.FitToHeight = int(p)
	}
}

// getPageLayout provides a method to get the fit to height for the worksheet.
func (p *FitToHeight) getPageLayout(ps *xlsxPageSetUp) {
	if ps == nil || ps.FitToHeight == 0 {
		*p = 1
		return
	}
	*p = FitToHeight(ps.FitToHeight)
}

// setPageLayout provides a method to set the fit to width for the worksheet.
func (p FitToWidth) setPageLayout(ps *xlsxPageSetUp) {
	if int(p) > 0 {
		ps.FitToWidth = int(p)
	}
}

// getPageLayout provides a method to get the fit to width for the worksheet.
func (p *FitToWidth) getPageLayout(ps *xlsxPageSetUp) {
	if ps == nil || ps.FitToWidth == 0 {
		*p = 1
		return
	}
	*p = FitToWidth(ps.FitToWidth)
}

// setPageLayout provides a method to set the scale for the worksheet.
func (p PageLayoutScale) setPageLayout(ps *xlsxPageSetUp) {
	if 10 <= int(p) && int(p) <= 400 {
		ps.Scale = int(p)
	}
}

// getPageLayout provides a method to get the scale for the worksheet.
func (p *PageLayoutScale) getPageLayout(ps *xlsxPageSetUp) {
	if ps == nil || ps.Scale < 10 || ps.Scale > 400 {
		*p = 100
		return
	}
	*p = PageLayoutScale(ps.Scale)
}

// SetPageLayout provides a function to sets worksheet page layout.
//
// Available options:
//
//    BlackAndWhite(bool)
//    FirstPageNumber(uint)
//    PageLayoutOrientation(string)
//    PageLayoutPaperSize(int)
//    FitToHeight(int)
//    FitToWidth(int)
//    PageLayoutScale(uint)
//
// The following shows the paper size sorted by xlsx index number:
//
//     Index | Paper Size
//    -------+-----------------------------------------------
//       1   | Letter paper (8.5 in. by 11 in.)
//       2   | Letter small paper (8.5 in. by 11 in.)
//       3   | Tabloid paper (11 in. by 17 in.)
//       4   | Ledger paper (17 in. by 11 in.)
//       5   | Legal paper (8.5 in. by 14 in.)
//       6   | Statement paper (5.5 in. by 8.5 in.)
//       7   | Executive paper (7.25 in. by 10.5 in.)
//       8   | A3 paper (297 mm by 420 mm)
//       9   | A4 paper (210 mm by 297 mm)
//       10  | A4 small paper (210 mm by 297 mm)
//       11  | A5 paper (148 mm by 210 mm)
//       12  | B4 paper (250 mm by 353 mm)
//       13  | B5 paper (176 mm by 250 mm)
//       14  | Folio paper (8.5 in. by 13 in.)
//       15  | Quarto paper (215 mm by 275 mm)
//       16  | Standard paper (10 in. by 14 in.)
//       17  | Standard paper (11 in. by 17 in.)
//       18  | Note paper (8.5 in. by 11 in.)
//       19  | #9 envelope (3.875 in. by 8.875 in.)
//       20  | #10 envelope (4.125 in. by 9.5 in.)
//       21  | #11 envelope (4.5 in. by 10.375 in.)
//       22  | #12 envelope (4.75 in. by 11 in.)
//       23  | #14 envelope (5 in. by 11.5 in.)
//       24  | C paper (17 in. by 22 in.)
//       25  | D paper (22 in. by 34 in.)
//       26  | E paper (34 in. by 44 in.)
//       27  | DL envelope (110 mm by 220 mm)
//       28  | C5 envelope (162 mm by 229 mm)
//       29  | C3 envelope (324 mm by 458 mm)
//       30  | C4 envelope (229 mm by 324 mm)
//       31  | C6 envelope (114 mm by 162 mm)
//       32  | C65 envelope (114 mm by 229 mm)
//       33  | B4 envelope (250 mm by 353 mm)
//       34  | B5 envelope (176 mm by 250 mm)
//       35  | B6 envelope (176 mm by 125 mm)
//       36  | Italy envelope (110 mm by 230 mm)
//       37  | Monarch envelope (3.875 in. by 7.5 in.).
//       38  | 6 3/4 envelope (3.625 in. by 6.5 in.)
//       39  | US standard fanfold (14.875 in. by 11 in.)
//       40  | German standard fanfold (8.5 in. by 12 in.)
//       41  | German legal fanfold (8.5 in. by 13 in.)
//       42  | ISO B4 (250 mm by 353 mm)
//       43  | Japanese postcard (100 mm by 148 mm)
//       44  | Standard paper (9 in. by 11 in.)
//       45  | Standard paper (10 in. by 11 in.)
//       46  | Standard paper (15 in. by 11 in.)
//       47  | Invite envelope (220 mm by 220 mm)
//       50  | Letter extra paper (9.275 in. by 12 in.)
//       51  | Legal extra paper (9.275 in. by 15 in.)
//       52  | Tabloid extra paper (11.69 in. by 18 in.)
//       53  | A4 extra paper (236 mm by 322 mm)
//       54  | Letter transverse paper (8.275 in. by 11 in.)
//       55  | A4 transverse paper (210 mm by 297 mm)
//       56  | Letter extra transverse paper (9.275 in. by 12 in.)
//       57  | SuperA/SuperA/A4 paper (227 mm by 356 mm)
//       58  | SuperB/SuperB/A3 paper (305 mm by 487 mm)
//       59  | Letter plus paper (8.5 in. by 12.69 in.)
//       60  | A4 plus paper (210 mm by 330 mm)
//       61  | A5 transverse paper (148 mm by 210 mm)
//       62  | JIS B5 transverse paper (182 mm by 257 mm)
//       63  | A3 extra paper (322 mm by 445 mm)
//       64  | A5 extra paper (174 mm by 235 mm)
//       65  | ISO B5 extra paper (201 mm by 276 mm)
//       66  | A2 paper (420 mm by 594 mm)
//       67  | A3 transverse paper (297 mm by 420 mm)
//       68  | A3 extra transverse paper (322 mm by 445 mm)
//       69  | Japanese Double Postcard (200 mm x 148 mm)
//       70  | A6 (105 mm x 148 mm)
//       71  | Japanese Envelope Kaku #2
//       72  | Japanese Envelope Kaku #3
//       73  | Japanese Envelope Chou #3
//       74  | Japanese Envelope Chou #4
//       75  | Letter Rotated (11in x 8 1/2 11 in)
//       76  | A3 Rotated (420 mm x 297 mm)
//       77  | A4 Rotated (297 mm x 210 mm)
//       78  | A5 Rotated (210 mm x 148 mm)
//       79  | B4 (JIS) Rotated (364 mm x 257 mm)
//       80  | B5 (JIS) Rotated (257 mm x 182 mm)
//       81  | Japanese Postcard Rotated (148 mm x 100 mm)
//       82  | Double Japanese Postcard Rotated (148 mm x 200 mm)
//       83  | A6 Rotated (148 mm x 105 mm)
//       84  | Japanese Envelope Kaku #2 Rotated
//       85  | Japanese Envelope Kaku #3 Rotated
//       86  | Japanese Envelope Chou #3 Rotated
//       87  | Japanese Envelope Chou #4 Rotated
//       88  | B6 (JIS) (128 mm x 182 mm)
//       89  | B6 (JIS) Rotated (182 mm x 128 mm)
//       90  | (12 in x 11 in)
//       91  | Japanese Envelope You #4
//       92  | Japanese Envelope You #4 Rotated
//       93  | PRC 16K (146 mm x 215 mm)
//       94  | PRC 32K (97 mm x 151 mm)
//       95  | PRC 32K(Big) (97 mm x 151 mm)
//       96  | PRC Envelope #1 (102 mm x 165 mm)
//       97  | PRC Envelope #2 (102 mm x 176 mm)
//       98  | PRC Envelope #3 (125 mm x 176 mm)
//       99  | PRC Envelope #4 (110 mm x 208 mm)
//       100 | PRC Envelope #5 (110 mm x 220 mm)
//       101 | PRC Envelope #6 (120 mm x 230 mm)
//       102 | PRC Envelope #7 (160 mm x 230 mm)
//       103 | PRC Envelope #8 (120 mm x 309 mm)
//       104 | PRC Envelope #9 (229 mm x 324 mm)
//       105 | PRC Envelope #10 (324 mm x 458 mm)
//       106 | PRC 16K Rotated
//       107 | PRC 32K Rotated
//       108 | PRC 32K(Big) Rotated
//       109 | PRC Envelope #1 Rotated (165 mm x 102 mm)
//       110 | PRC Envelope #2 Rotated (176 mm x 102 mm)
//       111 | PRC Envelope #3 Rotated (176 mm x 125 mm)
//       112 | PRC Envelope #4 Rotated (208 mm x 110 mm)
//       113 | PRC Envelope #5 Rotated (220 mm x 110 mm)
//       114 | PRC Envelope #6 Rotated (230 mm x 120 mm)
//       115 | PRC Envelope #7 Rotated (230 mm x 160 mm)
//       116 | PRC Envelope #8 Rotated (309 mm x 120 mm)
//       117 | PRC Envelope #9 Rotated (324 mm x 229 mm)
//       118 | PRC Envelope #10 Rotated (458 mm x 324 mm)
//
func (f *File) SetPageLayout(sheet string, opts ...PageLayoutOption) error {
	s, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	ps := s.PageSetUp
	if ps == nil {
		ps = new(xlsxPageSetUp)
		s.PageSetUp = ps
	}

	for _, opt := range opts {
		opt.setPageLayout(ps)
	}
	return err
}

// GetPageLayout provides a function to gets worksheet page layout.
//
// Available options:
//   PageLayoutOrientation(string)
//   PageLayoutPaperSize(int)
//   FitToHeight(int)
//   FitToWidth(int)
func (f *File) GetPageLayout(sheet string, opts ...PageLayoutOptionPtr) error {
	s, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	ps := s.PageSetUp

	for _, opt := range opts {
		opt.getPageLayout(ps)
	}
	return err
}

// SetDefinedName provides a function to set the defined names of the workbook
// or worksheet. If not specified scope, the default scope is workbook.
// For example:
//
//    f.SetDefinedName(&xlsx.DefinedName{
//        Name:     "Amount",
//        RefersTo: "Sheet1!$A$2:$D$5",
//        Comment:  "defined name comment",
//        Scope:    "Sheet2",
//    })
//
func (f *File) SetDefinedName(definedName *DefinedName) error {
	wb := f.workbookReader()
	d := xlsxDefinedName{
		Name:    definedName.Name,
		Comment: definedName.Comment,
		Data:    definedName.RefersTo,
	}
	if definedName.Scope != "" {
		if sheetIndex := f.GetSheetIndex(definedName.Scope); sheetIndex >= 0 {
			d.LocalSheetID = &sheetIndex
		}
	}
	if wb.DefinedNames != nil {
		for _, dn := range wb.DefinedNames.DefinedName {
			var scope string
			if dn.LocalSheetID != nil {
				scope = f.GetSheetName(*dn.LocalSheetID)
			}
			if scope == definedName.Scope && dn.Name == definedName.Name {
				return ErrDefinedNameduplicate
			}
		}
		wb.DefinedNames.DefinedName = append(wb.DefinedNames.DefinedName, d)
		return nil
	}
	wb.DefinedNames = &xlsxDefinedNames{
		DefinedName: []xlsxDefinedName{d},
	}
	return nil
}

// DeleteDefinedName provides a function to delete the defined names of the
// workbook or worksheet. If not specified scope, the default scope is
// workbook. For example:
//
//    f.DeleteDefinedName(&xlsx.DefinedName{
//        Name:     "Amount",
//        Scope:    "Sheet2",
//    })
//
func (f *File) DeleteDefinedName(definedName *DefinedName) error {
	wb := f.workbookReader()
	if wb.DefinedNames != nil {
		for idx, dn := range wb.DefinedNames.DefinedName {
			scope := "Workbook"
			deleteScope := definedName.Scope
			if deleteScope == "" {
				deleteScope = "Workbook"
			}
			if dn.LocalSheetID != nil {
				scope = f.GetSheetName(*dn.LocalSheetID)
			}
			if scope == deleteScope && dn.Name == definedName.Name {
				wb.DefinedNames.DefinedName = append(wb.DefinedNames.DefinedName[:idx], wb.DefinedNames.DefinedName[idx+1:]...)
				return nil
			}
		}
	}
	return ErrDefinedNameScope
}

// GetDefinedName provides a function to get the defined names of the workbook
// or worksheet.
func (f *File) GetDefinedName() []DefinedName {
	var definedNames []DefinedName
	wb := f.workbookReader()
	if wb.DefinedNames != nil {
		for _, dn := range wb.DefinedNames.DefinedName {
			definedName := DefinedName{
				Name:     dn.Name,
				Comment:  dn.Comment,
				RefersTo: dn.Data,
				Scope:    "Workbook",
			}
			if dn.LocalSheetID != nil && *dn.LocalSheetID >= 0 {
				definedName.Scope = f.GetSheetName(*dn.LocalSheetID)
			}
			definedNames = append(definedNames, definedName)
		}
	}
	return definedNames
}

// GroupSheets provides a function to group worksheets by given worksheets
// name. Group worksheets must contain an active worksheet.
func (f *File) GroupSheets(sheets []string) error {
	// check an active worksheet in group worksheets
	var inActiveSheet bool
	activeSheet := f.GetActiveSheetIndex()
	sheetMap := f.GetSheetList()
	for idx, sheetName := range sheetMap {
		for _, s := range sheets {
			if s == sheetName && idx == activeSheet {
				inActiveSheet = true
			}
		}
	}
	if !inActiveSheet {
		return ErrGroupSheets
	}
	// check worksheet exists
	wss := []*xlsxWorksheet{}
	for _, sheet := range sheets {
		worksheet, err := f.workSheetReader(sheet)
		if err != nil {
			return err
		}
		wss = append(wss, worksheet)
	}
	for _, ws := range wss {
		sheetViews := ws.SheetViews.SheetView
		if len(sheetViews) > 0 {
			for idx := range sheetViews {
				ws.SheetViews.SheetView[idx].TabSelected = true
			}
			continue
		}
	}
	return nil
}

// UngroupSheets provides a function to ungroup worksheets.
func (f *File) UngroupSheets() error {
	activeSheet := f.GetActiveSheetIndex()
	for index, sheet := range f.GetSheetList() {
		if activeSheet == index {
			continue
		}
		ws, _ := f.workSheetReader(sheet)
		sheetViews := ws.SheetViews.SheetView
		if len(sheetViews) > 0 {
			for idx := range sheetViews {
				ws.SheetViews.SheetView[idx].TabSelected = false
			}
		}
	}
	return nil
}

// InsertPageBreak create a page break to determine where the printed page
// ends and where begins the next one by given worksheet name and axis, so the
// content before the page break will be printed on one page and after the
// page break on another.
func (f *File) InsertPageBreak(sheet, cell string) (err error) {
	var ws *xlsxWorksheet
	var row, col int
	var rowBrk, colBrk = -1, -1
	if ws, err = f.workSheetReader(sheet); err != nil {
		return
	}
	if col, row, err = CellNameToCoordinates(cell); err != nil {
		return
	}
	col--
	row--
	if col == row && col == 0 {
		return
	}
	if ws.RowBreaks == nil {
		ws.RowBreaks = &xlsxBreaks{}
	}
	if ws.ColBreaks == nil {
		ws.ColBreaks = &xlsxBreaks{}
	}

	for idx, brk := range ws.RowBreaks.Brk {
		if brk.ID == row {
			rowBrk = idx
		}
	}
	for idx, brk := range ws.ColBreaks.Brk {
		if brk.ID == col {
			colBrk = idx
		}
	}

	if row != 0 && rowBrk == -1 {
		ws.RowBreaks.Brk = append(ws.RowBreaks.Brk, &xlsxBrk{
			ID:  row,
			Max: 16383,
			Man: true,
		})
		ws.RowBreaks.ManualBreakCount++
	}
	if col != 0 && colBrk == -1 {
		ws.ColBreaks.Brk = append(ws.ColBreaks.Brk, &xlsxBrk{
			ID:  col,
			Max: 1048575,
			Man: true,
		})
		ws.ColBreaks.ManualBreakCount++
	}
	ws.RowBreaks.Count = len(ws.RowBreaks.Brk)
	ws.ColBreaks.Count = len(ws.ColBreaks.Brk)
	return
}

// RemovePageBreak remove a page break by given worksheet name and axis.
func (f *File) RemovePageBreak(sheet, cell string) (err error) {
	var ws *xlsxWorksheet
	var row, col int
	if ws, err = f.workSheetReader(sheet); err != nil {
		return
	}
	if col, row, err = CellNameToCoordinates(cell); err != nil {
		return
	}
	col--
	row--
	if col == row && col == 0 {
		return
	}
	removeBrk := func(ID int, brks []*xlsxBrk) []*xlsxBrk {
		for i, brk := range brks {
			if brk.ID == ID {
				brks = append(brks[:i], brks[i+1:]...)
			}
		}
		return brks
	}
	if ws.RowBreaks == nil || ws.ColBreaks == nil {
		return
	}
	rowBrks := len(ws.RowBreaks.Brk)
	colBrks := len(ws.ColBreaks.Brk)
	if rowBrks > 0 && rowBrks == colBrks {
		ws.RowBreaks.Brk = removeBrk(row, ws.RowBreaks.Brk)
		ws.ColBreaks.Brk = removeBrk(col, ws.ColBreaks.Brk)
		ws.RowBreaks.Count = len(ws.RowBreaks.Brk)
		ws.ColBreaks.Count = len(ws.ColBreaks.Brk)
		ws.RowBreaks.ManualBreakCount--
		ws.ColBreaks.ManualBreakCount--
		return
	}
	if rowBrks > 0 && rowBrks > colBrks {
		ws.RowBreaks.Brk = removeBrk(row, ws.RowBreaks.Brk)
		ws.RowBreaks.Count = len(ws.RowBreaks.Brk)
		ws.RowBreaks.ManualBreakCount--
		return
	}
	if colBrks > 0 && colBrks > rowBrks {
		ws.ColBreaks.Brk = removeBrk(col, ws.ColBreaks.Brk)
		ws.ColBreaks.Count = len(ws.ColBreaks.Brk)
		ws.ColBreaks.ManualBreakCount--
	}
	return
}

// relsReader provides a function to get the pointer to the structure
// after deserialization of xl/worksheets/_rels/sheet%d.xml.rels.
func (f *File) relsReader(path string) *xlsxRelationships {
	var err error
	rels, _ := f.Relationships.Load(path)
	if rels == nil {
		if _, ok := f.Pkg.Load(path); ok {
			c := xlsxRelationships{}
			if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(path)))).
				Decode(&c); err != nil && err != io.EOF {
				log.Printf("xml decode error: %s", err)
			}
			f.Relationships.Store(path, &c)
		}
	}
	if rels, _ = f.Relationships.Load(path); rels != nil {
		return rels.(*xlsxRelationships)
	}
	return nil
}

// fillSheetData ensures there are enough rows, and columns in the chosen
// row to accept data. Missing rows are backfilled and given their row number
// Uses the last populated row as a hint for the size of the next row to add
func prepareSheetXML(ws *xlsxWorksheet, col int, row int) {
	ws.Lock()
	defer ws.Unlock()
	rowCount := len(ws.SheetData.Row)
	sizeHint := 0
	var ht float64
	var customHeight bool
	if ws.SheetFormatPr != nil && ws.SheetFormatPr.CustomHeight {
		ht = ws.SheetFormatPr.DefaultRowHeight
		customHeight = true
	}
	if rowCount > 0 {
		sizeHint = len(ws.SheetData.Row[rowCount-1].C)
	}
	if rowCount < row {
		// append missing rows
		for rowIdx := rowCount; rowIdx < row; rowIdx++ {
			ws.SheetData.Row = append(ws.SheetData.Row, xlsxRow{R: rowIdx + 1, CustomHeight: customHeight, Ht: ht, C: make([]xlsxC, 0, sizeHint)})
		}
	}
	rowData := &ws.SheetData.Row[row-1]
	fillColumns(rowData, col, row)
}

// fillColumns fill cells in the column of the row as contiguous.
func fillColumns(rowData *xlsxRow, col, row int) {
	cellCount := len(rowData.C)
	if cellCount < col {
		for colIdx := cellCount; colIdx < col; colIdx++ {
			cellName, _ := CoordinatesToCellName(colIdx+1, row)
			rowData.C = append(rowData.C, xlsxC{R: cellName})
		}
	}
}

// makeContiguousColumns make columns in specific rows as contiguous.
func makeContiguousColumns(ws *xlsxWorksheet, fromRow, toRow, colCount int) {
	ws.Lock()
	defer ws.Unlock()
	for ; fromRow < toRow; fromRow++ {
		rowData := &ws.SheetData.Row[fromRow-1]
		fillColumns(rowData, colCount, fromRow)
	}
}
