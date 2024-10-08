<?php

use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Shared\Converter;

include_once 'Sample_Header.php';

// New Word document
echo date('H:i:s'), ' Create new PhpWord object', EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();

// Define styles
$figureCaptionStyle = 'figureCaptionStyle';
$phpWord->addParagraphStyle($figureCaptionStyle, ['spaceAfter' => 120, 'spaceBefore' => 0, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'next' => 'Normal']);

// Begin code
$section = $phpWord->addSection();
$section->addText('Local image with alt text but no styles:');
$section->addImage('resources/_mars.jpg', null, 'Image of Mars');
$section->addCaption('Figure', 'Image of Mars', ['bold' => true], $figureCaptionStyle);

printSeparator($section);
$section->addText('Local image with styles but not alt text:');
$section->addImage('resources/_earth.jpg', ['width' => 210, 'height' => 210, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER]);
$section->addCaption('Figure', 'Image of Earth', ['bold' => true], $figureCaptionStyle);

// Remote image
printSeparator($section);
$source = 'http://php.net/images/logos/php-med-trans-light.gif';
$section->addText("Remote image from: {$source}");
$section->addImage($source, null);
$section->addCaption('Figure', 'Remote image', ['bold' => true], $figureCaptionStyle);

// Image from string
printSeparator($section);
$section->addText('In order to keep the image with the caption that follows, add the image to a textRun.');
$section->addTextBreak();
$source = 'resources/_mars.jpg';
$fileContent = file_get_contents($source);
$section->addText('Image from string');
$imageTextRun = $section->addTextRun(['keepNext' => true]);
$imageTextRun->addImage($fileContent);
$section->addCaption('Figure', 'Image of Mars', ['bold' => true], $figureCaptionStyle);

//Wrapping style
printSeparator($section);
$text = str_repeat('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. ', 2);
$wrappingStyles = ['inline', 'behind', 'infront', 'square', 'tight'];
foreach ($wrappingStyles as $wrappingStyle) {
    $section->addText("Wrapping style {$wrappingStyle}");
    $section->addImage(
        'resources/_earth.jpg',
        [
            'positioning' => 'relative',
            'marginTop' => -1,
            'marginLeft' => 1,
            'width' => 80,
            'height' => 80,
            'wrappingStyle' => $wrappingStyle,
            'wrapDistanceRight' => Converter::cmToPoint(1),
            'wrapDistanceBottom' => Converter::cmToPoint(1),
        ]
    );
    $section->addText($text);
    printSeparator($section);
}

//Absolute positioning
$section->addText('Absolute positioning: see top right corner of page');
$section->addImage(
    'resources/_mars.jpg',
    [
        'width' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(3),
        'height' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(3),
        'positioning' => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
        'posHorizontal' => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_RIGHT,
        'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_PAGE,
        'posVerticalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_PAGE,
        'marginLeft' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(15.5),
        'marginTop' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(1.55),
    ]
);

//Relative positioning
printSeparator($section);
$section->addText('Relative positioning: Horizontal position center relative to column,');
$section->addText('Vertical position top relative to line');
$section->addImage(
    'resources/_mars.jpg',
    [
        'width' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(3),
        'height' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(3),
        'positioning' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE,
        'posHorizontal' => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_CENTER,
        'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_COLUMN,
        'posVertical' => \PhpOffice\PhpWord\Style\Image::POSITION_VERTICAL_TOP,
        'posVerticalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_LINE,
    ]
);

function printSeparator(Section $section): void
{
    $section->addTextBreak();
    $lineStyle = ['weight' => 0.2, 'width' => 150, 'height' => 0, 'align' => 'center'];
    $section->addLine($lineStyle);
    $section->addTextBreak(2);
}

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
