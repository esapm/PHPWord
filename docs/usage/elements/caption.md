# Caption

To add a caption, use the ``addCaption`` method to sections, headers, footers, textruns, or table cells.

``` php
<?php

$section->addCaption($text, $label, [$fontStyle], [$paragraphStyle]);
```

- ``$text``. String describing the element being captioned.
- ``$label``. String describing the type of element being captioned (e.g., 'Figure' or 'Table').
- ``$fontStyle``: See [`Styles > Font`](../styles/font.md).
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
    'Image of Mars', 
    'Figure', 
    ['bold'=>true], 
    ['spaceAfter'=>240, 'spaceBefore'=>0]
);
```