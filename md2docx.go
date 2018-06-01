package md2docx // import "github.com/aerissecure/md2docx"

import (
	"errors"
	"fmt"
	"io"
	"strings"

	bf "gopkg.in/russross/blackfriday.v2"

	"baliance.com/gooxml/color"
	"baliance.com/gooxml/document"
	// "baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
)

// list formatting should be correct once this is fixed:
// https://github.com/baliance/gooxml/issues/136#issuecomment-367218322

var (
	gray = color.RGB(242, 242, 242)
)

// DocxRendererParameters configuration object that gets passed
// to NewDocxRenderer.
type DocxRendererParameters struct {
	StyleHyperlink     string
	StyleListOrdered   string
	StyleListUnordered string
	StyleHeading1      string
	StyleHeading2      string
	StyleHeading3      string
	StyleHeading4      string
	StyleHeading5      string
	StyleCodeBlock     string
	StyleCodeInline    string
	StyleBlockQuote    string
	StyleTable         string // This is a table specific style
}

// DocxRenderer is a type that implements the Renderer interface for docx output.
//
// Do not create this directly, instead use the NewDocxRenderer function.
type DocxRenderer struct {
	DocxRendererParameters

	doc       *document.Document
	para      document.Paragraph
	table     document.Table
	row       document.Row
	cell      document.Cell
	tableHead bool
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
		doc:       doc,
		listLevel: -1, // -1: not in a list, 0: first level of list
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
	// fmt.Println(node.Type)

	switch node.Type {

	case bf.Code: // gets text
		// we don't support line breaks in code. line breaks in code do not
		// trigger hardbreak, so don't use returns between ``

		// when in a table, case bf.Text runs before this with an empty Literal. bug?
		run := r.para.AddRun()
		run.AddText(string(node.Literal))
		run.Properties().SetStyle(r.StyleCodeInline)

	case bf.CodeBlock: // gets text, is not child of paragraph
		// info is the text on the same line as the opening ~~~, such as programming language ~~~python
		// i could use this to controls some type of formatting maybe
		r.para = r.doc.AddParagraph()
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
		r.para = r.doc.AddParagraph()
		style := r.getHeading(node.HeadingData.Level)
		r.para.Properties().SetStyle(style)

	case bf.BlockQuote:
		if !entering {
			r.para.Properties().SetStyle(r.StyleBlockQuote)
		}

	case bf.Paragraph:
		if !entering {
			break
		}
		r.para = r.doc.AddParagraph()
		if node.Parent.Type == bf.Item {
			r.setListStyle(r.para, r.listLevel, node.Parent.Parent.ListData.ListFlags)
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

		if node.Parent.Type == bf.TableCell {
			lines := strings.Split(string(node.Literal), "\\n")

			run := r.para.AddRun()
			if r.strong {
				run.Properties().SetBold(true)
			}
			if r.emph {
				run.Properties().SetItalic(true)
			}

			run.AddText(lines[0])

			if len(lines) > 1 {
				for _, line := range lines[1:] {
					// new para and duplicate existing para properties
					para := r.cell.AddParagraph()
					para.Properties().X().Jc = r.para.Properties().X().Jc
					r.para = para

					run := r.para.AddRun()
					if r.strong {
						run.Properties().SetBold(true)
					}
					if r.emph {
						run.Properties().SetItalic(true)
					}

					run.AddText(line)
				}
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

	case bf.Table:
		if !entering {
			break
		}
		r.table = r.doc.AddTable()
		r.table.Properties().SetWidthPercent(100)
		if r.StyleTable != "" {
			r.table.Properties().SetStyle(r.StyleTable)
		}

	case bf.TableHead:
		// not currently being used
		r.tableHead = entering

	case bf.TableBody:
		break

	case bf.TableRow:
		if !entering {
			break
		}
		r.row = r.table.AddRow()

	case bf.TableCell:
		if !entering {
			break
		}
		r.cell = r.row.AddCell()
		r.para = r.cell.AddParagraph()
		align := cellAlignment(node.Align)
		r.para.Properties().SetAlignment(align)
		// If i keep the special handling for Cell under bf.Text, then i may
		// not want to create the paragraph here
		// for multiple paragraphs in same cell, is there a way to just copy the
		// properties of the previous paragraph??
		// Using \n as a newline in tables is a good idea i think. That is obviously
		// not how it would work in html, but i think its ok for input text to render
		// slightly differently based on the output. having \n mean something in docx
		// output and not in html output is better than using <br> for docx output.
		// Also word and html treat returns differently, in html they get ignored and
		// word they get displayed as a newline or paragraph
		//
		// para.Properties().X().CT_PPr
		//
		// an escaped backslack is actually being used, its a \n literal, not a newline
		// that we are matching on.

	// handled but uneeded node types:
	case bf.Document, bf.Link, bf.Item:
		break

	default:
		fmt.Println("unsupported node type", node.Type.String())
		// probably default to just putting it into a regular paragraph?

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

// setListStyle styles a paragraph as a list. to do so requires
// setting (1) paragraph style, (2) indentation level, (3) numbering definition
func (r *DocxRenderer) setListStyle(para document.Paragraph, lvl int, flags bf.ListType) {
	if lvl < 0 {
		// don't set style if indent level is not set
		return
	}
	ordered := flags&bf.ListTypeOrdered != 0
	var style string
	if ordered {
		style = r.StyleListOrdered
	} else {
		style = r.StyleListUnordered
	}
	// style is set on all indentation levels
	para.SetStyle(style)
	// ilvl & numId only set starting the second indentation
	if lvl > 0 {
		nd, err := r.styleToNumDef(style)
		if err != nil {
			fmt.Println("error retrieving numbering definition", err)
		}
		r.para.SetNumberingLevel(lvl)
		r.para.SetNumberingDefinition(nd)
	}
}

// styleToNumDef gets the NumberingDefinition that is used to format a list
// for the style. The style is typically referenced in the first lvl element
// under abstractNum. However, it is unclear if the style will always be
// referenced from the abstractNum.
func (r *DocxRenderer) styleToNumDef(style string) (document.NumberingDefinition, error) {
	var errNumDef document.NumberingDefinition

	numDefs := r.doc.Numbering.Definitions()
	if len(numDefs) == 0 {
		return errNumDef, errors.New("no numbering definitions found")
	}
	for _, nd := range numDefs {
		for _, lvl := range nd.X().Lvl {
			if lvl.PStyle != nil && style == lvl.PStyle.ValAttr {
				return nd, nil
			}
		}
	}
	return errNumDef, fmt.Errorf("numbering definition not found for style: %s", style)
}

// cellAlignment gets the docx alignment object indicated by
// the CellAlignFlags.
func cellAlignment(align bf.CellAlignFlags) wml.ST_Jc {
	switch align {
	case bf.TableAlignmentLeft:
		return wml.ST_JcLeft
	case bf.TableAlignmentCenter:
		return wml.ST_JcCenter
	case bf.TableAlignmentRight:
		return wml.ST_JcRight
	default:
		return wml.ST_JcUnset
	}
}
