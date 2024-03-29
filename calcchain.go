package xlsx

import (
	"bytes"
	"encoding/xml"
	"io"
	"log"
)

// calcChainReader provides a function to get the pointer to the structure
// after deserialization of xl/calcChain.xml.
func (f *File) calcChainReader() *xlsxCalcChain {
	var err error

	if f.CalcChain == nil {
		f.CalcChain = new(xlsxCalcChain)
		if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML("xl/calcChain.xml")))).
			Decode(f.CalcChain); err != nil && err != io.EOF {
			log.Printf("xml decode error: %s", err)
		}
	}

	return f.CalcChain
}

// calcChainWriter provides a function to save xl/calcChain.xml after
// serialize structure.
func (f *File) calcChainWriter() {
	if f.CalcChain != nil && f.CalcChain.C != nil {
		output, _ := xml.Marshal(f.CalcChain)
		f.saveFileList("xl/calcChain.xml", output)
	}
}

// deleteCalcChain provides a function to remove cell reference on the
// calculation chain.
func (f *File) deleteCalcChain(index int, axis string) {
	calc := f.calcChainReader()
	if calc != nil {
		calc.C = xlsxCalcChainCollection(calc.C).Filter(func(c xlsxCalcChainC) bool {
			return !((c.I == index && c.R == axis) || (c.I == index && axis == ""))
		})
	}
	if len(calc.C) == 0 {
		f.CalcChain = nil
		f.Pkg.Delete("xl/calcChain.xml")
		content := f.contentTypesReader()
		content.Lock()
		defer content.Unlock()
		for k, v := range content.Overrides {
			if v.PartName == "/xl/calcChain.xml" {
				content.Overrides = append(content.Overrides[:k], content.Overrides[k+1:]...)
			}
		}
	}
}

type xlsxCalcChainCollection []xlsxCalcChainC

// Filter provides a function to filter calculation chain.
func (c xlsxCalcChainCollection) Filter(fn func(v xlsxCalcChainC) bool) []xlsxCalcChainC {
	var results []xlsxCalcChainC
	for _, v := range c {
		if fn(v) {
			results = append(results, v)
		}
	}
	return results
}
