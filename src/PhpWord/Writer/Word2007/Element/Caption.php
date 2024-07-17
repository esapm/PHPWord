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

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

use PhpOffice\PhpWord\Style;

/**
 * TextRun element writer.
 *
 * @since 0.10.0
 */
class Caption extends AbstractElement
{
    /**
     * Write caption element.
     */
    public function write(): void
    {
        $xmlWriter = $this->getXmlWriter();
        $element = $this->getElement();
        if (!$element instanceof \PhpOffice\PhpWord\Element\Caption) {
            return;
        }

        $this->startElementP();

        $bookmarkRId = null;
        $rId = $element->getRelationId();
        $bookmarkRId = $element->getPhpWord()->addBookmark();
        $figureNumber = $element->getFigureNumber();
        $label = $element->getLabel();

        // Bookmark start for Table of Figures
        $xmlWriter->startElement('w:bookmarkStart');
        $xmlWriter->writeAttribute('w:id', $bookmarkRId);
        $xmlWriter->writeAttribute('w:name', '_' . $label . (string) $rId);
        $xmlWriter->endElement(); //w:bookmarkStart

        // Label
        $xmlWriter->startElement('w:r');

        $this->writeFontStyle();

        $xmlWriter->startElement('w:t');
        $xmlWriter->writeAttribute('xml:space', 'preserve');
        $this->writeText($label . ' ');
        $xmlWriter->endElement(); // w:t
        $xmlWriter->endElement(); // w:r

        // Section heading (future). The text value increments for every relevant section. 
        // The period is automatic and separates the section number from the figure number.
        // <w:r><w:instrText xml:space="preserve"> STYLEREF 1 \s </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>.</w:t></w:r>

        // Sequence
        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'begin');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:instrText');
        $xmlWriter->writeAttribute('xml:space', 'preserve');
        $xmlWriter->text(" SEQ {$label} \* ARABIC ");
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'separate');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');

        $this->writeFontStyle();

        $xmlWriter->startElement('w:rPr');
        $xmlWriter->startElement('w:noProof');
        $xmlWriter->endElement();
        $xmlWriter->endElement();
        $xmlWriter->startElement('w:t');
        $xmlWriter->text($figureNumber);
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'end');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        // Actual text
        $text = $element->getText();
        if (is_string($text)) {
            $xmlWriter->startElement('w:r');

            $this->writeFontStyle();

            $xmlWriter->startElement('w:t');
            $this->writeText('. ' . $text);
            $xmlWriter->endElement(); // w:t
            $xmlWriter->endElement(); // w:r
        }

        // Bookmark end
        $xmlWriter->startElement('w:bookmarkEnd');
        $xmlWriter->writeAttribute('w:id', $bookmarkRId);
        $xmlWriter->endElement(); //w:bookmarkEnd

        $this->endElementP(); // w:p
    }
}
