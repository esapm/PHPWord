# Caption

To add a caption, use the ``addCaption`` method to sections, headers, footers, textruns, or table cells.

``` php
<?php

$section->addCaption($label, [$text], [$fontStyle], [$paragraphStyle]);
```

- ``$label``. String describing the type of element being captioned (e.g., 'Figure' or 'Table').
- ``$text``. String describing the element being captioned.
- ``$fontStyle``: See [`Styles > Font`](../styles/font.md). Applies only to the caption text (not the label or number). It overrides the fontStyle associated with $paragraphStyle.
- ``$paragraphStyle``: See [`Styles > Paragraph`](../styles/paragraph.md).

Examples:

``` php
<?php

$section = $phpWord->addSection();
$section->addImage(
    'mars.jpg',
    array(
        'width'         => 100,
        'height'        => 100,
        'marginTop'     => -1,
        'marginLeft'    => -1,
        'wrappingStyle' => 'behind',
        'keepNext'      => true
    ),
    'alt text'
);
$section->addCaption(
    'Figure', 
    'Image of Mars', 
    ['bold'=>false], 
    ['spaceAfter'=>240, 'spaceBefore'=>0, 'fontStyle' => array('bold' => true)]
);
```