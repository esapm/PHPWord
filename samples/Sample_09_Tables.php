<?php

use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\TablePosition;

include_once 'Sample_Header.php';

// New Word Document
echo date('H:i:s'), ' Create new PhpWord object', EOL;
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$section = $phpWord->addSection();
$header = ['size' => 16, 'bold' => true];

// 1. Basic table

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

// 2. Advanced table

$section->addTextBreak(1);
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
$table->addTblCaption('This is a table caption. It will appear in the Title field of the Alt Text Properties tab.');
$table->addTblDescription('This is a table description. It will appear in the Description field of the Alt Text Properties tab.');

/*
 *  3. colspan (gridSpan) and rowspan (vMerge)
 *  ---------------------
 *  |     |   B    |    |
 *  |  A  |--------|  E |
 *  |     | C |  D |    |
 *  ---------------------
 */

$section->addPageBreak();
$section->addText('Table with colspan and rowspan', $header);

$fancyTableStyle = ['borderSize' => 6, 'borderColor' => '999999'];
$cellRowSpan = ['vMerge' => 'restart', 'valign' => 'center', 'bgColor' => 'FFFF00'];
$cellRowContinue = ['vMerge' => 'continue'];
$cellColSpan = ['gridSpan' => 2, 'valign' => 'center'];
$cellHCentered = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER];
$cellVCentered = ['valign' => 'center'];

$spanTableStyleName = 'Colspan Rowspan';
$phpWord->addTableStyle($spanTableStyleName, $fancyTableStyle);
$section->addCaption('Colspan (gridSpan) and rowspan (vMerge)', 'Table', $fancyTableFontStyle, $fancyTableCaptionStyle);
$table = $section->addTable($spanTableStyleName);
// Note: Alt text (table captions and descriptions) will be ignored if added to a named style.
// Use table->addTblCaption() and table->addTblDescription() instead.

$table->addRow();

$cell1 = $table->addCell(2000, $cellRowSpan);
$textrun1 = $cell1->addTextRun($cellHCentered);
$textrun1->addText('A');
$textrun1->addFootnote()->addText('Row span');

$cell2 = $table->addCell(4000, $cellColSpan);
$textrun2 = $cell2->addTextRun($cellHCentered);
$textrun2->addText('B');
$textrun2->addFootnote()->addText('Column span');

$table->addCell(2000, $cellRowSpan)->addText('E', null, $cellHCentered);

$table->addRow();
$table->addCell(null, $cellRowContinue);
$table->addCell(2000, $cellVCentered)->addText('C', null, $cellHCentered);
$table->addCell(2000, $cellVCentered)->addText('D', null, $cellHCentered);
$table->addCell(null, $cellRowContinue);

/*
 *  4. colspan (gridSpan) and rowspan (vMerge)
 *  ---------------------
 *  |     |   B    |  1 |
 *  |  A  |        |----|
 *  |     |        |  2 |
 *  |     |---|----|----|
 *  |     | C |  D |  3 |
 *  ---------------------
 * @see https://github.com/PHPOffice/PHPWord/issues/806
 */

$section->addPageBreak();
$section->addText('Table with colspan and rowspan', $header);

$styleTable = ['borderSize' => 6, 'borderColor' => '999999', 'tblCaption' => 'Table captions can be added in the style array.', 'tblDescription' => 'Table descriptions can also be added in the style array. Both will be overwritten by text entered using the addTblCaption() or addTblDescription() functions.'];
$phpWord->addTableStyle('Colspan Rowspan', $styleTable);
$section->addCaption('Colspan (gridSpan) and rowspan (vMerge)', 'Table', $fancyTableFontStyle, $fancyTableCaptionStyle);
$table = $section->addTable('Colspan Rowspan');

$row = $table->addRow();
$row->addCell(1000, ['vMerge' => 'restart'])->addText('A');
$row->addCell(1000, ['gridSpan' => 2, 'vMerge' => 'restart'])->addText('B');
$row->addCell(1000)->addText('1');

$row = $table->addRow();
$row->addCell(1000, ['vMerge' => 'continue']);
$row->addCell(1000, ['vMerge' => 'continue', 'gridSpan' => 2]);
$row->addCell(1000)->addText('2');

$row = $table->addRow();
$row->addCell(1000, ['vMerge' => 'continue']);
$row->addCell(1000)->addText('C');
$row->addCell(1000)->addText('D');
$row->addCell(1000)->addText('3');

// 5. Nested table

$section->addTextBreak(2);
$section->addText('Nested table in a centered and 50% width table.', $header);

$table = $section->addTable(['width' => 50 * 50, 'unit' => 'pct', 'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER]);
$cell = $table->addRow()->addCell();
$cell->addText('This cell contains nested table.');
$innerCell = $cell->addTable(['alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER])->addRow()->addCell();
$innerCell->addText('Inside nested table');

// 6. Table with floating position

$section->addTextBreak(2);
$section->addText('Table with floating positioning.', $header);

$table = $section->addTable(['borderSize' => 6, 'borderColor' => '999999', 'position' => ['vertAnchor' => TablePosition::VANCHOR_TEXT, 'bottomFromText' => Converter::cmToTwip(1)]]);
$cell = $table->addRow()->addCell();
$cell->addText('This is a single cell.');

// Save file
echo write($phpWord, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
