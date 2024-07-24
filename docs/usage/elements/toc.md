# Table of contents

To add a table of contents (TOC), you can use the ``addTOC`` method.
Your TOC can only be generated if you have add at least one title (See "[Title](title.md)").

``` php
<?php

$section->addTOC([$fontStyle], [$tocStyle], [$minDepth], [$maxDepth], [$switchString], [$tField]);
```

- ``$fontStyle``. See font style section.
- ``$tocStyle``. See available options below.
- ``$minDepth``. Minimum depth of header to be shown. Default 1.
- ``$maxDepth``. Maximum depth of header to be shown. Default 9.
- ``$switchString``. Switches to define TOC. Default '\h \z \u'. See below and ISO/IEC 29500-1.
- ``$tField``. Heading, depth pairs added to the TOC if the \t switch is set.

Options for ``$tocStyle``:

- ``tabLeader``. Fill type between the title text and the page number. Use the defined constants in ``\PhpOffice\PhpWord\Style\TOC``.
- ``tabPos``. The position of the tab where the page number appears in *twip*.
- ``indent``. The indent factor of the titles in *twip*.

## Notes

- Like Word, PHPWord will automatically assign style TOC X to TOC entries, where X is the depth or paragraph outline level. Therefore, it is strongly recommended that TOC X styles be defined. When doing so, the 'customStyle' array element must be set to '0'.

- There are three primary ways to create a TOC.
  1. With \o switch and heading depths.
  2. With \t switch and a $switchString containing "Style,depth,Style,depth" pairs. This switch can be used in conjunction with the \o switch.
  3. With the \u switch. This will build a TOC using the outline levels associated with title styles. Outline levels are assigned by PHPWord via numbering levels ($numLevel). (Currently, a $numStyle and $numLevel must be applied even if the style level will not have a number.) The \u switch can conflict with the \t switch if a style in the $tField string does not have an assigned outline level ($numLevel).