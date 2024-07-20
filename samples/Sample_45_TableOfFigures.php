<?php

use PhpOffice\PhpWord\Element\Section;


include_once 'Sample_Header.php';

// New Word document
echo date('H:i:s'), ' Create new PhpWord object', EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$phpWord->getSettings()->setUpdateFields(true);

// New section
$section = $phpWord->addSection();

// Define styles
$fontStyle12 = ['spaceAfter' => 60, 'size' => 12];
$fontStyle10 = ['size' => 10];
$header = ['size' => 16, 'bold' => true];
$phpWord->addTitleStyle(null, ['size' => 22, 'bold' => true]);
$figureCaptionStyle = 'figureCaptionStyle';
$phpWord->addParagraphStyle($figureCaptionStyle, ['spaceAfter' => 120, 'spaceBefore' => 0, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START,'next' => 'Normal']);

// Modify the default or built-in "Table of Figures" style. This will ensure that the table maintains the associated formatting even if the entire table is updated within Word.
$tofStyle = 'Table of Figures';
$phpWord->addParagraphStyle($tofStyle, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::START, 'spaceAfter' => 100, 'customStyle' => '0']);

// Add text elements
$section->addTitle('Table of figures', 0);

// Add Table of figures
$tof = $section->addTOF('Figure', ['tabLeader' => PhpOffice\PhpWord\Style\TOC::TAB_LEADER_DOT, 'tabPos' => 9360], null, $tofStyle);
$section->addTextBreak(1);

// Add text elements
$section->addTitle('Table of tables', 0);

// Add Table of tables
$toc = $section->addTOF('Table', ['tabLeader' => PhpOffice\PhpWord\Style\TOC::TAB_LEADER_DOT, 'tabPos' => 9360], $fontStyle12, $tofStyle);
$section->addTextBreak(1);

// Add images
$section->addText('Image 1:');
$section->addImage('resources/_mars.jpg', null, 'Image of Mars');
$section->addCaption('Image of Mars', 'Figure', ['bold' => true], $figureCaptionStyle);

printSeparator($section);
$section->addText('Local image with styles but not alt text:');
$section->addImage('resources/_earth.jpg', ['width' => 210, 'height' => 210, 'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER]);
$section->addCaption('Image of Earth', 'Figure', ['bold' => true], $figureCaptionStyle);

// Remote image
printSeparator($section);
$source = 'http://php.net/images/logos/php-med-trans-light.gif';
$section->addText("Remote image from: {$source}");
$section->addImage($source, null);
$section->addCaption('Remote image', 'Figure', ['bold' => true], $figureCaptionStyle);

printSeparator($section);

// Basic table

$rows = 10;
$cols = 5;
$section->addText('Basic table', $header);

$section->addCaption('Basic table', 'Table', ['bold' => true], ['spaceAfter' => 0, 'spaceBefore' => 240, 'keepNext' => true]);
$table = $section->addTable();
for ($r = 1; $r <= $rows; ++$r) {
    $table->addRow();
    for ($c = 1; $c <= $cols; ++$c) {
        $table->addCell(1750)->addText("Row {$r}, Cell {$c}");
    }
}

// Advanced table

$section->addTextBreak(1);

printSeparator($section);
$section->addText('Fancy table', $header);

$fancyTableStyleName = 'Fancy Table';
$fancyTableStyle = ['borderSize' => 6, 'borderColor' => '006699', 'cellMargin' => 80, 'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER, 'cellSpacing' => 50];
$fancyTableFirstRowStyle = ['borderBottomSize' => 18, 'borderBottomColor' => '0000FF', 'bgColor' => '66BBFF'];
$fancyTableCellStyle = ['valign' => 'center'];
$fancyTableCellBtlrStyle = ['valign' => 'center', 'textDirection' => \PhpOffice\PhpWord\Style\Cell::TEXT_DIR_BTLR];
$fancyTableFontStyle = ['bold' => true];
$fancyTableCaptionStyle = ['spaceAfter' => 0, 'spaceBefore' => 240, 'keepNext' => true];
$phpWord->addTableStyle($fancyTableStyleName, $fancyTableStyle, $fancyTableFirstRowStyle);

$section->addCaption('Fancy table', 'Table', $fancyTableFontStyle, $fancyTableCaptionStyle);

$table = $section->addTable($fancyTableStyleName);
$table->addRow(900);
$table->addCell(2000, $fancyTableCellStyle)->addText('Row 1', $fancyTableFontStyle);
$table->addCell(2000, $fancyTableCellStyle)->addText('Row 2', $fancyTableFontStyle);
$table->addCell(2000, $fancyTableCellStyle)->addText('Row 3', $fancyTableFontStyle);
$table->addCell(2000, $fancyTableCellStyle)->addText('Row 4', $fancyTableFontStyle);
$table->addCell(500, $fancyTableCellBtlrStyle)->addText('Row 5', $fancyTableFontStyle);
for ($i = 1; $i <= 8; ++$i) {
    $table->addRow();
    $table->addCell(2000)->addText("Cell {$i}");
    $table->addCell(2000)->addText("Cell {$i}");
    $table->addCell(2000)->addText("Cell {$i}");
    $table->addCell(2000)->addText("Cell {$i}");
    $text = (0 == $i % 2) ? 'X' : '';
    $table->addCell(500)->addText($text);
}

echo date('H:i:s'), ' Note: Please refresh TOC manually.', EOL;

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
