package xlsx

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddComments(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	s := strings.Repeat("c", 32768)
	assert.NoError(t, f.AddComment("Sheet1", "A30", s, s))
	assert.NoError(t, f.AddComment("Sheet2", "B7", "xlsx: ", "This is a comment."))

	// Test add comment on not exists worksheet.
	assert.EqualError(t, f.AddComment("SheetN", "B7", "xlsx: ", "This is a comment."), "sheet SheetN is not exist")
	// Test add comment on with illegal cell coordinates
	assert.EqualError(t, f.AddComment("Sheet1", "A", "xlsx: ", "This is a comment."), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	if assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddComments.xlsx"))) {
		assert.Len(t, f.GetComments(), 2)
	}

	f.Comments["xl/comments2.xml"] = nil
	f.Pkg.Store("xl/comments2.xml", []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>xlsx: </author></authors><commentList><comment ref="B7" authorId="0"><text><t>xlsx: </t></text></comment></commentList></comments>`))
	comments := f.GetComments()
	assert.EqualValues(t, 2, len(comments["Sheet1"]))
	assert.EqualValues(t, 1, len(comments["Sheet2"]))
	assert.EqualValues(t, len(NewFile().GetComments()), 0)
}

func TestDecodeVMLDrawingReader(t *testing.T) {
	f := NewFile()
	path := "xl/drawings/vmlDrawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	f.decodeVMLDrawingReader(path)
}

func TestCommentsReader(t *testing.T) {
	f := NewFile()
	path := "xl/comments1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	f.commentsReader(path)
}

func TestCountComments(t *testing.T) {
	f := NewFile()
	f.Comments["xl/comments1.xml"] = nil
	assert.Equal(t, f.countComments(), 1)
}
