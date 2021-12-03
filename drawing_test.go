package xlsx

import (
	"sync"
	"testing"
)

func TestDrawingParser(t *testing.T) {
	f := File{
		Drawings: sync.Map{},
		Pkg:      sync.Map{},
	}
	f.Pkg.Store("charset", MacintoshCyrillicCharset)
	f.Pkg.Store("wsDr", []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:oneCellAnchor><xdr:graphicFrame/></xdr:oneCellAnchor></xdr:wsDr>`))
	// Test with one cell anchor
	f.drawingParser("wsDr")
	// Test with unsupported charset
	f.drawingParser("charset")
}
