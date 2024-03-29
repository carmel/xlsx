package xlsx

import "encoding/xml"

// xlsxProperties specifies to an OOXML document properties such as the
// template used, the number of pages and words, and the application name and
// version.
type xlsxProperties struct {
	XMLName              xml.Name `xml:"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties Properties"`
	Template             string
	Manager              string
	Company              string
	Pages                int
	Words                int
	Characters           int
	PresentationFormat   string
	Lines                int
	Paragraphs           int
	Slides               int
	Notes                int
	TotalTime            int
	HiddenSlides         int
	MMClips              int
	ScaleCrop            bool
	HeadingPairs         *xlsxVectorVariant
	TitlesOfParts        *xlsxVectorLpstr
	LinksUpToDate        bool
	CharactersWithSpaces int
	SharedDoc            bool
	HyperlinkBase        string
	HLinks               *xlsxVectorVariant
	HyperlinksChanged    bool
	DigSig               *xlsxDigSig
	Application          string
	AppVersion           string
	DocSecurity          int
}

// xlsxVectorVariant specifies the set of hyperlinks that were in this
// document when last saved.
type xlsxVectorVariant struct {
	Content string `xml:",innerxml"`
}

type xlsxVectorLpstr struct {
	Content string `xml:",innerxml"`
}

// xlsxDigSig contains the signature of a digitally signed document.
type xlsxDigSig struct {
	Content string `xml:",innerxml"`
}
