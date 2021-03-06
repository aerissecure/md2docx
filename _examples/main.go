package main

import (
	"fmt"

	"baliance.com/gooxml/document"
	bf "gopkg.in/russross/blackfriday.v2"

	"github.com/aerissecure/md2docx"
)

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

> block quote here
> second line of block quote
> third line of block quote?

Regular paragraph with ` + "`code code code`" + ` inside

code span with line break like this` + " `code code\ncode` " + `should
actually just render as a single space, which words because docx doesn't
interpret \\n and just renders it as a space anyway.

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

lets try a bulleted list with multiple paragraphs in a bullet:

- first item

    Continuing the first item
- second item

    continuing second item

        $ with this whatever it is
        $ with this again

followed by the following ordered:

1. this is num 1
1. this is num 2
1. this is num 3
    1. this is indented a level _**and bold/ital**_
    1. this is indented a level

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

Last paragraph, for **_bold italic_**.

# Tables Section


|Key | Value|
|:----:|------|
|hey | ` + "`you`" + `|
|cell2 | *hey*|
|cell3 | _hey_ | you
|**1** | **2**|
|new \ line| test|
|- new \n - line| - test\n- test2|
|--|--|
|asdf|asdf|
| | asdf|
| | asdf |
`)

func main() {
	// doc := document.New()
	doc, _ := document.OpenTemplate("template.docx")
	for _, s := range doc.Styles.Styles() {
		fmt.Println("style", s.Name(), "has ID of", s.StyleID(), "type is", s.Type())
	}

	params := md2docx.DocxRendererParameters{
		StyleHyperlink:     "Hyperlink",
		StyleListOrdered:   "ListOrdered",
		StyleListUnordered: "ListParagraph",
		StyleHeading1:      "Heading1",
		StyleHeading2:      "Heading2",
		StyleHeading3:      "Heading3",
		StyleHeading4:      "Heading4",
		StyleHeading5:      "Heading5",
		StyleCodeBlock:     "IntenseQuote",
		StyleCodeInline:    "BookTitle",
		StyleBlockQuote:    "Quote",
		StyleTable:         "GridTable4-Accent1",
	}

	// Much of what we do will be rendering into a table cell or specific location
	// in a document. What do we want to accept as being passed in?
	// it would be cool to just be able to render without a document, but I don't
	// think that works.
	renderer := md2docx.NewDocxRenderer(doc, params)
	bf.Run(input, bf.WithRenderer(renderer), bf.WithExtensions(bf.CommonExtensions))
	doc.SaveToFile("out.docx")
}
