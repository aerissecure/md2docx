package md2docx

import (
	"io"
	"strings"

	"baliance.com/gooxml/document"
	"baliance.com/gooxml/schema/soo/wml"
	bf "gopkg.in/russross/blackfriday.v2"
)

// DocxRendererParameters configuration object that gets passed
// to NewDocxRenderer.
type DocxRendererParameters struct {
	StyleHyperlink    string
	StyleListOrdered  string
	StyleListBulleted string
	StyleHeading1     string
	StyleHeading2     string
	StyleHeading3     string
	StyleHeading4     string
	StyleHeading5     string
	StyleCodeBlock    string
	StyleCodeInline   string
}

// DocxRenderer is a type that implements the Renderer interface for docx output.
//
// Do not create this directly, instead use the NewDocxRenderer function.
type DocxRenderer struct {
	DocxRendererParameters

	Document  *document.Document
	para      document.Paragraph
	listLevel int
	strong    bool
	emph      bool
}

func (r *DocxRenderer) run() document.Run {
	runs := r.para.Runs()
	return runs[len(runs)-1] // should always exist
}

func (r *DocxRenderer) getHeading(level int) string {
	switch level {
	case 1:
		return r.StyleHeading1
	case 2:
		return r.StyleHeading2
	case 3:
		return r.StyleHeading3
	case 4:
		return r.StyleHeading4
	default:
		return r.StyleHeading5
	}
}

// NewDocxRenderer creates and configures a DocxRenderer object, which
// satisfies the Renderer interface.
func NewDocxRenderer(doc *document.Document, params DocxRendererParameters) *DocxRenderer {
	// params must specify all styles that will be used. Default is "", which defaults to Normal
	return &DocxRenderer{
		DocxRendererParameters: params,
		Document:               doc,
		listLevel:              -1, // -1: not in a list, 0: first level of list
	}
}

// RenderNode is a default renderer of a single node of a syntax tree. For
// block nodes it will be called twice: first time with entering=true, second
// time with entering=false, so that it could know when it's working on an open
// tag and when on close. It writes the result to w.
//
// The return value is a way to tell the calling walker to adjust its walk
// pattern: e.g. it can terminate the traversal by returning Terminate. Or it
// can ask the walker to skip a subtree of this node by returning SkipChildren.
// The typical behavior is to return GoToNext, which asks for the usual
// traversal to the next node.
func (r *DocxRenderer) RenderNode(w io.Writer, node *bf.Node, entering bool) bf.WalkStatus {

	switch node.Type {

	case bf.Code: // gets text
		// we don't support line breaks in code. line breaks in code do not
		// trigger hardbreak, so don't use returns between ``
		run := r.para.AddRun()
		run.AddText(string(node.Literal))
		run.Properties().SetStyle(r.StyleCodeInline)

	case bf.CodeBlock: // gets text, is not child of paragraph
		// info is the text on the same line as the opening ~~~, such as programming language ~~~python
		// i could use this to controls some type of formatting maybe
		r.para = r.Document.AddParagraph()
		r.para.Properties().SetStyle(r.StyleCodeBlock)
		run := r.para.AddRun()
		// HardLine does not get called inside of code blocks

		lines := strings.Split(string(node.Literal), "\n")
		for i, line := range lines {
			run.AddText(line)
			// no break after last two tokens (2nd to last is last text, last is an additional break from leaving the block)
			if i+2 < len(lines) {
				run.AddBreak()
			}
		}

	case bf.Heading: // no paragraph child
		if !entering {
			break
		}
		r.para = r.Document.AddParagraph()
		style := r.getHeading(node.HeadingData.Level)
		r.para.Properties().SetStyle(style)

	case bf.Paragraph:
		if !entering {
			break
		}
		r.para = r.Document.AddParagraph()
		if node.Parent.Type == bf.Item {
			r.setListStyle(r.para, node.Parent.Parent.ListData.ListFlags)
			if r.listLevel > 0 {
				numpr := wml.NewCT_NumPr()
				lvl := wml.NewCT_DecimalNumber()
				lvl.ValAttr = int64(r.listLevel)
				numpr.Ilvl = lvl
				numpr.NumId = lvl
				r.para.X().PPr.NumPr = numpr
			}
			// do something with r.listLevel
		}

	// Softbreak: a simple single newline("\n"). In html can be rendered as a space or a carriage return.
	// https://github.com/russross/blackfriday/blob/v2/inline.go#L162
	// case bf.Softbreak: // simple new line ("\n"). this is not implemented and will never trigger: https://github.com/russross/blackfriday/issues/315

	// Hardbreak: a single newline/return preceded by at least two spaces. translated as a <br>
	case bf.Hardbreak:
		// using HardLineBreak extension has unintended affect on list Items
		// a simple list will count the returns between items as hard breaks
		// and rendering them leads to a list with spaces.
		// if HardLineBreak is used, this can be overcome by checking node.Parent.Parent == Item
		r.run().AddBreak()

	case bf.Strong:
		r.strong = entering
	case bf.Emph:
		r.emph = entering
	case bf.Text:
		if node.Parent.Type == bf.Link {
			linkText := string(node.Literal)
			altText := string(node.Parent.Title)
			dest := string(node.Parent.Destination)
			hl := r.para.AddHyperLink()
			hl.SetTarget(dest)
			run := hl.AddRun()
			run.Properties().SetStyle(r.StyleHyperlink)
			run.AddText(linkText)
			if altText != "" {
				hl.SetToolTip(altText)
			}
			break
		}

		run := r.para.AddRun()
		if r.strong {
			run.Properties().SetBold(true)
		}
		if r.emph {
			run.Properties().SetItalic(true)
		}
		run.AddText(string(node.Literal))

	case bf.List:
		if entering {
			r.listLevel++
		} else {
			r.listLevel--
		}

	// handled but uneeded node types:
	case bf.Document, bf.Link, bf.Item:
		break
	default:
		panic("unsupported node type: " + node.Type.String())
	}
	return bf.GoToNext
}

// RenderHeader writes document preamble and TOC if requested.
func (r *DocxRenderer) RenderHeader(w io.Writer, ast *bf.Node) {
	return
	// io.WriteString(w, "header is written here\n")
}

// RenderFooter writes document footer.
func (r *DocxRenderer) RenderFooter(w io.Writer, ast *bf.Node) {
	return
	// io.WriteString(w, "footer is written here")
}

func (r *DocxRenderer) setListStyle(para document.Paragraph, flags bf.ListType) {
	ordered := flags&bf.ListTypeOrdered != 0
	if ordered {
		para.Properties().SetStyle(r.StyleListOrdered)
	} else {
		para.Properties().SetStyle(r.StyleListBulleted)
	}
}
