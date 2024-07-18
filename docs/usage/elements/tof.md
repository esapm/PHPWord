# Table of figures

To add a table of figures (TOF), you can use the ``addTOF`` method.
Your TOF can only be generated if you have add at least one caption with a label (See "[Caption](caption.md)").

``` php
<?php

$section->addTOF($captionLabel, [$tofStyle], [$fontStyle], [$paragraphStyle]);
```

- ``$captionLabel``. The value of the caption labels to include (e.g., 'Figure', 'Table', etcl).
- ``$tofStyle``. See available options below.
- ``$fontStyle``. See font style section.
- ``$paragraphStyle``. See [`Styles > Paragraph`](../styles/paragraph.md).

Options for ``$tofStyle``:

- ``tabLeader``. Fill type between the title text and the page number. Use the defined constants in ``\PhpOffice\PhpWord\Style\TOC``.
- ``tabPos``. The position of the tab where the page number appears in *twip*.