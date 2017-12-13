//
// Blackfriday Markdown Processor
// Available at http://github.com/russross/blackfriday
//
// Copyright © 2011 Russ Ross <russ@russross.com>.
// Distributed under the Simplified BSD License.
// See README.md for details.
//

//
//
// HTML rendering backend
//
//

package md2docx

import (
	"io"
	"strings"

	"baliance.com/gooxml/document"
	bf "gopkg.in/russross/blackfriday.v2"
)

// DocxRendererParameters
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

	Document *document.Document
	para     document.Paragraph
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
		run := r.para.AddRun()
		run.AddText(string(node.Literal))
		run.Properties().SetStyle(r.StyleCodeInline)

	case bf.CodeBlock: // gets text, is not child of paragraph
		// info is the text on the same line as the opening ~~~, such as programming language ~~~python
		// i could use this to controls some type of formatting maybe
		r.para = r.Document.AddParagraph()
		r.para.Properties().SetStyle(r.StyleCodeBlock)
		run := r.para.AddRun()
		breakingRun(run, string(node.Literal))

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
		}

	// case bf.Softbreak: // simple new line ("\n"). this is not implemented and will never trigger: https://github.com/russross/blackfriday/issues/315

	// case bf.Hardbreak: // new line with two spaces preceding ("  \n", or two new lines ("\n\n")

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
		if node.Parent.Type == bf.Strong {
			run.Properties().SetBold(true)
		}
		if node.Parent.Type == bf.Emph {
			run.Properties().SetItalic(true)
		}
		breakingRun(run, string(node.Literal))

	// handled but uneeded node types:
	case bf.Document, bf.Strong, bf.Emph, bf.List, bf.Item, bf.Link:
		break
	default:
		panic("unsupported node type: " + node.Type.String())
	}
	return bf.GoToNext
}

// // RenderHeader writes HTML document preamble and TOC if requested.
func (r *DocxRenderer) RenderHeader(w io.Writer, ast *bf.Node) {
	return
	// io.WriteString(w, "header is written here\n")
}

// // RenderFooter writes HTML document footer.
func (r *DocxRenderer) RenderFooter(w io.Writer, ast *bf.Node) {
	return
	// io.WriteString(w, "footer is written here")
}

// breakingRun renders a run that has line breaks (soft breaks), but no
// other markdown tokens.
func breakingRun(run document.Run, text string) {
	lines := strings.Split(text, "\n")
	for i, line := range lines {
		run.AddText(line)
		// add break unless on last token
		if i+1 < len(lines) {
			run.AddBreak()
		}
	}
}

func (r *DocxRenderer) setListStyle(para document.Paragraph, flags bf.ListType) {
	ordered := flags&bf.ListTypeOrdered == 1
	if ordered {
		para.Properties().SetStyle(r.StyleListOrdered)
	} else {
		para.Properties().SetStyle(r.StyleListBulleted)
	}
}
