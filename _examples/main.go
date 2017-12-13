package main

import (
	// "fmt"

	"github.com/aerissecure/md2docx"
	bf "gopkg.in/russross/blackfriday.v2"

	"baliance.com/gooxml/document"
)

// soft break is just a single return. it can have

var input = []byte(`
lets try a soft
break

` + "lets try a hard  " + `
break

this is **bold** and *italic* font
not a new paragraph

should be a new paragraph
soft
.

    Adding a new paragraph that is code indented 4
    and a second line of code

Regular paragraph.

going to add more spacing, two lines


going to add more spacing, three lines



fenced block:

~~~
Fenced code block
and a second line of fenced code
~~~

Regular paragraph.
` + "\nRegular paragraph with `code code\n\ncode` inside" + `

Another paragraph, with the following:

- unordered first
- unordered second
- unordered last
    - this one is a second level of indent

a list with breaks

- first one, adding break here:\
am i still on the first bullet?
- second one, no break
- third one, three breaks\
break 2\
break 3\

followed by the following ordered:

1. this is num 1
1. this is num 2
1. this is num 3
    1. this is indentied a level

closing paragraph with a <http://google.com> link with no info

And a second closing paragraph with a [Title text here](https://www.google.com)

another para with [link text](https://aerissecure.com)

# h1

- is

- this

- still a list?

## h2

https://aerissecure.com should be parsed automatically

and now i'm going to try the\
backslack (this is new line)

### h3

`)

func main() {
	// doc := document.New()
	doc, _ := document.OpenTemplate("../template.docx")
	// for _, s := range doc.Styles.Styles() {
	// 	fmt.Println("style", s.Name(), "has ID of", s.StyleID(), "type is", s.Type())
	// }

	doc.AddParagraph().AddRun().AddText("\n\nasdfasdf\n\n\nasdfasdf\n\n")

	params := md2docx.DocxRendererParameters{
		StyleHyperlink:    "Hyperlink",
		StyleListOrdered:  "ListOrdered",
		StyleListBulleted: "ListParagraph",
		StyleHeading1:     "Heading1",
		StyleHeading2:     "Heading2",
		StyleHeading3:     "Heading3",
		StyleHeading4:     "Heading4",
		StyleHeading5:     "Heading5",
		StyleCodeBlock:    "IntenseQuote",
		StyleCodeInline:   "BookTitle",
	}

	// check with empty style

	renderer := md2docx.NewDocxRenderer(doc, params)
	bf.Run(input, bf.WithRenderer(renderer), bf.WithExtensions(bf.CommonExtensions))
	doc.SaveToFile("out.docx")
}
