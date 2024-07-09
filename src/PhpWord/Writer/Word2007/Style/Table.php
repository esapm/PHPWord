<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Style;

use PhpOffice\PhpWord\Shared\XMLWriter;
use PhpOffice\PhpWord\SimpleType\TblWidth;
use PhpOffice\PhpWord\Style\Table as TableStyle;
use PhpOffice\PhpWord\Writer\Word2007\Element\TableAlignment;

/**
 * Table style writer.
 *
 * @since 0.10.0
 */
class Table extends AbstractStyle
{
    /**
     * @var int Table width
     */
    private $width;

    /**
     * @var string tblCaption
     */
    private $tblCaption;

    /**
     * @var string tblDescription
     */
    private $tblDescription;

    /**
     * Write style.
     */
    public function write(): void
    {
        $style = $this->getStyle();
        $xmlWriter = $this->getXmlWriter();

        if ($style instanceof TableStyle) {
            $this->writeStyle($xmlWriter, $style);
        } elseif (is_string($style)) {
            $xmlWriter->startElement('w:tblPr');
            $xmlWriter->startElement('w:tblStyle');
            $xmlWriter->writeAttribute('w:val', $style);
            $xmlWriter->endElement();
            if (null !== $this->width) {
                $this->writeTblWidth($xmlWriter, 'w:tblW', TblWidth::PERCENT, $this->width);
            }
            // Even though Word groups alt text fields with table styles, the alt text is ignored when sent as part
            // of a named style. The code below allows table captions and descriptions to be added via the 
            // addTblCaption() and addTblDescription() functions. These populate the Alt Text Title and Alt Text
            // Description fields, respectively.
            //
            // Caption (alt text title)
            $this->writeTblCaption($xmlWriter, 'w:tblCaption', $this->tblCaption);

            // Description (alt text contents)
            $this->writeTblDescription($xmlWriter, 'w:tblDescription', $this->tblDescription);

            $xmlWriter->endElement();
        }
    }

    /**
     * Write full style.
     */
    private function writeStyle(XMLWriter $xmlWriter, TableStyle $style): void
    {
        // w:tblPr
        $xmlWriter->startElement('w:tblPr');

        // Table alignment
        if ('' !== $style->getAlignment()) {
            $tableAlignment = new TableAlignment($style->getAlignment());
            $xmlWriter->startElement($tableAlignment->getName());
            foreach ($tableAlignment->getAttributes() as $attributeName => $attributeValue) {
                $xmlWriter->writeAttribute($attributeName, $attributeValue);
            }
            $xmlWriter->endElement();
        }

        // If table alt text (captions and descriptions) variables are not set using the addTblCaption() and addTblDescription() functions,
        // the code below will include them if they are part of the table style.
        //
        // Caption (alt text title)
        $this->writeTblCaption($xmlWriter, 'w:tblCaption', (!empty($this->tblCaption) ? $this->tblCaption : $style->getTblCaption()));

        // Description (alt text contents)
        $this->writeTblDescription($xmlWriter, 'w:tblDescription', (!empty($this->tblDescription) ? $this->tblDescription : $style->getTblDescription()));

        $this->writeTblWidth($xmlWriter, 'w:tblW', $style->getUnit(), $style->getWidth());
        $this->writeTblWidth($xmlWriter, 'w:tblCellSpacing', TblWidth::TWIP, $style->getCellSpacing());
        $this->writeIndent($xmlWriter, $style);
        $this->writeLayout($xmlWriter, $style->getLayout());

        // Position
        $styleWriter = new TablePosition($xmlWriter, $style->getPosition());
        $styleWriter->write();

        //Right to left
        $xmlWriter->writeElementIf($style->isBidiVisual() !== null, 'w:bidiVisual', 'w:val', $this->writeOnOf($style->isBidiVisual()));

        $this->writeMargin($xmlWriter, $style);
        $this->writeBorder($xmlWriter, $style);

        $xmlWriter->endElement(); // w:tblPr

        $this->writeShading($xmlWriter, $style);

        // First row style
        $firstRow = $style->getFirstRow();
        if ($firstRow instanceof TableStyle) {
            $this->writeFirstRow($xmlWriter, $firstRow);
        }
    }

    /**
     * Enable/Disable automatic resizing of the table.
     *
     * @param string $layout autofit / fixed
     */
    private function writeLayout(XMLWriter $xmlWriter, $layout): void
    {
        $xmlWriter->startElement('w:tblLayout');
        $xmlWriter->writeAttribute('w:type', $layout);
        $xmlWriter->endElement(); // w:tblLayout
    }

    /**
     * Write margin.
     */
    private function writeMargin(XMLWriter $xmlWriter, TableStyle $style): void
    {
        if ($style->hasMargin()) {
            $xmlWriter->startElement('w:tblCellMar');

            $styleWriter = new MarginBorder($xmlWriter);
            $styleWriter->setSizes($style->getCellMargin());
            $styleWriter->write();

            $xmlWriter->endElement(); // w:tblCellMar
        }
    }

    /**
     * Write border.
     */
    private function writeBorder(XMLWriter $xmlWriter, TableStyle $style): void
    {
        if ($style->hasBorder()) {
            $xmlWriter->startElement('w:tblBorders');

            $styleWriter = new MarginBorder($xmlWriter);
            $styleWriter->setSizes($style->getBorderSize());
            $styleWriter->setColors($style->getBorderColor());
            $styleWriter->write();

            $xmlWriter->endElement(); // w:tblBorders
        }
    }

    /**
     * Writes a table width.
     *
     * @param string $elementName
     * @param string $unit
     * @param float|int $width
     */
    private function writeTblWidth(XMLWriter $xmlWriter, $elementName, $unit, $width = null): void
    {
        if (null === $width) {
            return;
        }
        $xmlWriter->startElement($elementName);
        $xmlWriter->writeAttributeIf(null !== $width, 'w:w', $width);
        $xmlWriter->writeAttribute('w:type', $unit);
        $xmlWriter->endElement();
    }

    /**
     * Write row style.
     */
    private function writeFirstRow(XMLWriter $xmlWriter, TableStyle $style): void
    {
        $xmlWriter->startElement('w:tblStylePr');
        $xmlWriter->writeAttribute('w:type', 'firstRow');
        $xmlWriter->startElement('w:tcPr');

        $this->writeBorder($xmlWriter, $style);
        $this->writeShading($xmlWriter, $style);

        $xmlWriter->endElement(); // w:tcPr
        $xmlWriter->endElement(); // w:tblStylePr
    }

    /**
     * Write shading.
     */
    private function writeShading(XMLWriter $xmlWriter, TableStyle $style): void
    {
        if (null !== $style->getShading()) {
            $xmlWriter->startElement('w:tcPr');

            $styleWriter = new Shading($xmlWriter, $style->getShading());
            $styleWriter->write();

            $xmlWriter->endElement();
        }
    }

    /**
     * Set width.
     *
     * @param int $value
     */
    public function setWidth($value = null): void
    {
        $this->width = $value;
    }

    /**
     * Set tblCaption.
     *
     * @param string $value
     */
    public function setTblCaption($value = null): void
    {
        $this->tblCaption = $value;
    }

    /**
     * Set tblDescription.
     *
     * @param string $value
     */
    public function setTblDescription($value = null): void
    {
        $this->tblDescription = $value;
    }

    private function writeIndent(XMLWriter $xmlWriter, TableStyle $style): void
    {
        $indent = $style->getIndent();

        if ($indent === null) {
            return;
        }

        $this->writeTblWidth($xmlWriter, 'w:tblInd', $indent->getType(), $indent->getValue());
    }

    /**
     * Write table caption (Title of alt text).
     *
     * @param string $elementName
     * @param string $tblCaption
     */
    private function writeTblCaption(XMLWriter $xmlWriter, $elementName, $tblCaption): void
    {
        if (empty($tblCaption)) {
            return;
        }
        $xmlWriter->startElement($elementName);
        $xmlWriter->writeAttribute('w:val', $tblCaption);
        $xmlWriter->endElement(); // w:tblCaption
    }

    /**
     * Write table description (alt text).
     *
     * @param string $elementName
     * @param string $tblDescription
     */
    private function writeTblDescription(XMLWriter $xmlWriter, $elementName, $tblDescription): void
    {
        if (empty($tblDescription)) {
            return;
        }
        $xmlWriter->startElement($elementName);
        $xmlWriter->writeAttribute('w:val', $tblDescription);
        $xmlWriter->endElement(); // w:tblDescription
    }
}
