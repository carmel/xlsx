package xlsx

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"path/filepath"
	"strconv"
	"strings"
)

// parseFormatCommentsSet provides a function to parse the format settings of
// the comment with default value.
func parseFormatCommentsSet(formatSet string) (*formatComment, error) {
	format := formatComment{
		Author: "Author:",
		Text:   " ",
	}
	err := json.Unmarshal([]byte(formatSet), &format)
	return &format, err
}

// GetComments retrieves all comments and returns a map of worksheet name to
// the worksheet comments.
func (f *File) GetComments() (comments map[string][]Comment) {
	comments = map[string][]Comment{}
	for n, path := range f.sheetMap {
		target := f.getSheetComments(filepath.Base(path))
		if target == "" {
			continue
		}
		if !strings.HasPrefix(target, "/") {
			target = "xl" + strings.TrimPrefix(target, "..")
		}
		if d := f.commentsReader(strings.TrimPrefix(target, "/")); d != nil {
			sheetComments := []Comment{}
			for _, comment := range d.CommentList.Comment {
				sheetComment := Comment{}
				if comment.AuthorID < len(d.Authors.Author) {
					sheetComment.Author = d.Authors.Author[comment.AuthorID]
				}
				sheetComment.Ref = comment.Ref
				sheetComment.AuthorID = comment.AuthorID
				if comment.Text.T != nil {
					sheetComment.Text += *comment.Text.T
				}
				for _, text := range comment.Text.R {
					if text.T != nil {
						sheetComment.Text += text.T.Val
					}
				}
				sheetComments = append(sheetComments, sheetComment)
			}
			comments[n] = sheetComments
		}
	}
	return
}

// getSheetComments provides the method to get the target comment reference by
// given worksheet file path.
func (f *File) getSheetComments(sheetFile string) string {
	var rels = "xl/worksheets/_rels/" + sheetFile + ".rels"
	if sheetRels := f.relsReader(rels); sheetRels != nil {
		sheetRels.Lock()
		defer sheetRels.Unlock()
		for _, v := range sheetRels.Relationships {
			if v.Type == SourceRelationshipComments {
				return v.Target
			}
		}
	}
	return ""
}

// AddComment provides the method to add comment in a sheet by given worksheet
// index, cell and format set (such as author and text). Note that the max
// author length is 255 and the max text length is 32512. For example, add a
// comment in Sheet1!$A$30:
//
//    err := f.AddComment("Sheet1", "A30", `{"author":"Excelize: ","text":"This is a comment."}`)
//
func (f *File) AddComment(sheet, cell, format string) error {
	formatSet, err := parseFormatCommentsSet(format)
	if err != nil {
		return err
	}
	// Read sheet data.
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	commentID := f.countComments() + 1
	drawingVML := "xl/drawings/vmlDrawing" + strconv.Itoa(commentID) + ".vml"
	sheetRelationshipsComments := "../comments" + strconv.Itoa(commentID) + ".xml"
	sheetRelationshipsDrawingVML := "../drawings/vmlDrawing" + strconv.Itoa(commentID) + ".vml"
	if ws.LegacyDrawing != nil {
		// The worksheet already has a comments relationships, use the relationships drawing ../drawings/vmlDrawing%d.vml.
		sheetRelationshipsDrawingVML = f.getSheetRelationshipsTargetByID(sheet, ws.LegacyDrawing.RID)
		commentID, _ = strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(sheetRelationshipsDrawingVML, "../drawings/vmlDrawing"), ".vml"))
		drawingVML = strings.Replace(sheetRelationshipsDrawingVML, "..", "xl", -1)
	} else {
		// Add first comment for given sheet.
		sheetRels := "xl/worksheets/_rels/" + strings.TrimPrefix(f.sheetMap[trimSheetName(sheet)], "xl/worksheets/") + ".rels"
		rID := f.addRels(sheetRels, SourceRelationshipDrawingVML, sheetRelationshipsDrawingVML, "")
		f.addRels(sheetRels, SourceRelationshipComments, sheetRelationshipsComments, "")
		f.addSheetNameSpace(sheet, SourceRelationship)
		f.addSheetLegacyDrawing(sheet, rID)
	}
	commentsXML := "xl/comments" + strconv.Itoa(commentID) + ".xml"
	var colCount int
	for i, l := range strings.Split(formatSet.Text, "\n") {
		if ll := len(l); ll > colCount {
			if i == 0 {
				ll += len(formatSet.Author)
			}
			colCount = ll
		}
	}
	err = f.addDrawingVML(commentID, drawingVML, cell, strings.Count(formatSet.Text, "\n")+1, colCount)
	if err != nil {
		return err
	}
	f.addComment(commentsXML, cell, formatSet)
	f.addContentTypePart(commentID, "comments")
	return err
}

// addDrawingVML provides a function to create comment as
// xl/drawings/vmlDrawing%d.vml by given commit ID and cell.
func (f *File) addDrawingVML(commentID int, drawingVML, cell string, lineCount, colCount int) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return err
	}
	yAxis := col - 1
	xAxis := row - 1
	vml := f.VMLDrawing[drawingVML]
	if vml == nil {
		vml = &vmlDrawing{
			XMLNSv:  "urn:schemas-microsoft-com:vml",
			XMLNSo:  "urn:schemas-microsoft-com:office:office",
			XMLNSx:  "urn:schemas-microsoft-com:office:excel",
			XMLNSmv: "http://macVmlSchemaUri",
			Shapelayout: &xlsxShapelayout{
				Ext: "edit",
				IDmap: &xlsxIDmap{
					Ext:  "edit",
					Data: commentID,
				},
			},
			Shapetype: &xlsxShapetype{
				ID:        "_x0000_t202",
				Coordsize: "21600,21600",
				Spt:       202,
				Path:      "m0,0l0,21600,21600,21600,21600,0xe",
				Stroke: &xlsxStroke{
					Joinstyle: "miter",
				},
				VPath: &vPath{
					Gradientshapeok: "t",
					Connecttype:     "rect",
				},
			},
		}
	}
	sp := encodeShape{
		Fill: &vFill{
			Color2: "#fbfe82",
			Angle:  -180,
			Type:   "gradient",
			Fill: &oFill{
				Ext:  "view",
				Type: "gradientUnscaled",
			},
		},
		Shadow: &vShadow{
			On:       "t",
			Color:    "black",
			Obscured: "t",
		},
		Path: &vPath{
			Connecttype: "none",
		},
		Textbox: &vTextbox{
			Style: "mso-direction-alt:auto",
			Div: &xlsxDiv{
				Style: "text-align:left",
			},
		},
		ClientData: &xClientData{
			ObjectType: "Note",
			Anchor: fmt.Sprintf(
				"%d, 23, %d, 0, %d, %d, %d, 5",
				1+yAxis, 1+xAxis, 2+yAxis+lineCount, colCount+yAxis, 2+xAxis+lineCount),
			AutoFill: "True",
			Row:      xAxis,
			Column:   yAxis,
		},
	}
	s, _ := xml.Marshal(sp)
	shape := xlsxShape{
		ID:          "_x0000_s1025",
		Type:        "#_x0000_t202",
		Style:       "position:absolute;73.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden",
		Fillcolor:   "#fbf6d6",
		Strokecolor: "#edeaa1",
		Val:         string(s[13 : len(s)-14]),
	}
	d := f.decodeVMLDrawingReader(drawingVML)
	if d != nil {
		for _, v := range d.Shape {
			s := xlsxShape{
				ID:          "_x0000_s1025",
				Type:        "#_x0000_t202",
				Style:       "position:absolute;73.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden",
				Fillcolor:   "#fbf6d6",
				Strokecolor: "#edeaa1",
				Val:         v.Val,
			}
			vml.Shape = append(vml.Shape, s)
		}
	}
	vml.Shape = append(vml.Shape, shape)
	f.VMLDrawing[drawingVML] = vml
	return err
}

// addComment provides a function to create chart as xl/comments%d.xml by
// given cell and format sets.
func (f *File) addComment(commentsXML, cell string, formatSet *formatComment) {
	a := formatSet.Author
	t := formatSet.Text
	if len(a) > MaxFieldLength {
		a = a[:MaxFieldLength]
	}
	if len(t) > 32512 {
		t = t[:32512]
	}
	comments := f.commentsReader(commentsXML)
	authorID := 0
	if comments == nil {
		comments = &xlsxComments{Authors: xlsxAuthor{Author: []string{formatSet.Author}}}
	}
	if inStrSlice(comments.Authors.Author, formatSet.Author) == -1 {
		comments.Authors.Author = append(comments.Authors.Author, formatSet.Author)
		authorID = len(comments.Authors.Author) - 1
	}
	defaultFont := f.GetDefaultFont()
	bold := ""
	cmt := xlsxComment{
		Ref:      cell,
		AuthorID: authorID,
		Text: xlsxText{
			R: []xlsxR{
				{
					RPr: &xlsxRPr{
						B:  &bold,
						Sz: &attrValFloat{Val: float64Ptr(9)},
						Color: &xlsxColor{
							Indexed: 81,
						},
						RFont:  &attrValString{Val: stringPtr(defaultFont)},
						Family: &attrValInt{Val: intPtr(2)},
					},
					T: &xlsxT{Val: a},
				},
				{
					RPr: &xlsxRPr{
						Sz: &attrValFloat{Val: float64Ptr(9)},
						Color: &xlsxColor{
							Indexed: 81,
						},
						RFont:  &attrValString{Val: stringPtr(defaultFont)},
						Family: &attrValInt{Val: intPtr(2)},
					},
					T: &xlsxT{Val: t},
				},
			},
		},
	}
	comments.CommentList.Comment = append(comments.CommentList.Comment, cmt)
	f.Comments[commentsXML] = comments
}

// countComments provides a function to get comments files count storage in
// the folder xl.
func (f *File) countComments() int {
	c1, c2 := 0, 0
	f.Pkg.Range(func(k, v interface{}) bool {
		if strings.Contains(k.(string), "xl/comments") {
			c1++
		}
		return true
	})
	for rel := range f.Comments {
		if strings.Contains(rel, "xl/comments") {
			c2++
		}
	}
	if c1 < c2 {
		return c2
	}
	return c1
}

// decodeVMLDrawingReader provides a function to get the pointer to the
// structure after deserialization of xl/drawings/vmlDrawing%d.xml.
func (f *File) decodeVMLDrawingReader(path string) *decodeVmlDrawing {
	var err error

	if f.DecodeVMLDrawing[path] == nil {
		c, ok := f.Pkg.Load(path)
		if ok && c != nil {
			f.DecodeVMLDrawing[path] = new(decodeVmlDrawing)
			if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(c.([]byte)))).
				Decode(f.DecodeVMLDrawing[path]); err != nil && err != io.EOF {
				log.Printf("xml decode error: %s", err)
			}
		}
	}
	return f.DecodeVMLDrawing[path]
}

// vmlDrawingWriter provides a function to save xl/drawings/vmlDrawing%d.xml
// after serialize structure.
func (f *File) vmlDrawingWriter() {
	for path, vml := range f.VMLDrawing {
		if vml != nil {
			v, _ := xml.Marshal(vml)
			f.Pkg.Store(path, v)
		}
	}
}

// commentsReader provides a function to get the pointer to the structure
// after deserialization of xl/comments%d.xml.
func (f *File) commentsReader(path string) *xlsxComments {
	var err error
	if f.Comments[path] == nil {
		content, ok := f.Pkg.Load(path)
		if ok && content != nil {
			f.Comments[path] = new(xlsxComments)
			if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(content.([]byte)))).
				Decode(f.Comments[path]); err != nil && err != io.EOF {
				log.Printf("xml decode error: %s", err)
			}
		}
	}
	return f.Comments[path]
}

// commentsWriter provides a function to save xl/comments%d.xml after
// serialize structure.
func (f *File) commentsWriter() {
	for path, c := range f.Comments {
		if c != nil {
			v, _ := xml.Marshal(c)
			f.saveFileList(path, v)
		}
	}
}
