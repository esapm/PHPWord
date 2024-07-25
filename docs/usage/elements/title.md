# Title

If you want to structure your document or build table of contents, you need titles or headings.
To add a title to the document, use the ``addTitleStyle`` and ``addTitle`` method.
If `depth` is 0, a Title will be inserted, otherwise a Heading1, Heading2, ...

``` php
<?php

$phpWord->addTitleStyle($depth, [$fontStyle], [$paragraphStyle], [$styleName]);
$section->addTitle($text, $depth, [$pageNumber], [$style]);
```

`addTitleStyle` :
- ``$depth`` : If no $styleName is given, $depth determines the Heading level (e.g., Heading 1, Heading 2, etc.). If a $styleName is given, its outline level will be set to $depth (assuming the paragraph style does not already exist). Note that the TOC \u switch can be used to build a TOC based on paragraph outline levels. Setting $depth = 10 will give it an outline level of "Body Text".
- ``$fontStyle``: See [`Styles > Font`](../styles/font.md).
- ``$paragraphStyle``: See [`Styles > Paragraph`](../styles/paragraph.md)
- ``$styleName``: Name of the heading style. Default 'Heading [$depth]'.

`addTitle`
- ``$text`` : Text to be displayed in the document. This can be `string` or a `\PhpOffice\PhpWord\Element\TextRun`
- ``$depth`` : If no $style value is given, $depth determines the Heading style to be applied to the title. Otherwise, $depth is ignored. Default 1.
- ``$pageNumber`` : Number of the page.
- ``$style`` : Name of the heading style. Default 'Heading [$depth]' or 'Title' if $depth == 0.

Paragraphs that do not have an associated title style cannot be added to the TOC.