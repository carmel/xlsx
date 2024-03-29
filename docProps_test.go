package xlsx

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

var MacintoshCyrillicCharset = []byte{0x8F, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2, 0x20, 0xEC, 0xE8, 0xF0}

func TestSetDocProps(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetDocProps(&DocProperties{
		Category:       "category",
		ContentStatus:  "Draft",
		Created:        "2019-06-04T22:00:10Z",
		Creator:        "Go Excelize",
		Description:    "This file created by Go Excelize",
		Identifier:     "xlsx",
		Keywords:       "Spreadsheet",
		LastModifiedBy: "Go Author",
		Modified:       "2019-06-04T22:00:10Z",
		Revision:       "0",
		Subject:        "Test Subject",
		Title:          "Test Title",
		Language:       "en-US",
		Version:        "1.0.0",
	}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetDocProps.xlsx")))
	f.Pkg.Store("docProps/core.xml", nil)
	assert.NoError(t, f.SetDocProps(&DocProperties{}))
	assert.NoError(t, f.Close())

	// Test unsupported charset
	f = NewFile()
	f.Pkg.Store("docProps/core.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetDocProps(&DocProperties{}), "xml decode error: XML syntax error on line 1: invalid UTF-8")
}

func TestGetDocProps(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	props, err := f.GetDocProps()
	assert.NoError(t, err)
	assert.Equal(t, props.Creator, "Microsoft Office User")
	f.Pkg.Store("docProps/core.xml", nil)
	_, err = f.GetDocProps()
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	// Test unsupported charset
	f = NewFile()
	f.Pkg.Store("docProps/core.xml", MacintoshCyrillicCharset)
	_, err = f.GetDocProps()
	assert.EqualError(t, err, "xml decode error: XML syntax error on line 1: invalid UTF-8")
}
