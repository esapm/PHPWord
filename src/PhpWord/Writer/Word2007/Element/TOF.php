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

use PhpOffice\PhpWord\Element\Caption;
use PhpOffice\PhpWord\Element\TOF as TOFElement;
use PhpOffice\PhpWord\Shared\XMLWriter;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Writer\Word2007\Style\Font as FontStyleWriter;
use PhpOffice\PhpWord\Writer\Word2007\Style\Paragraph as ParagraphStyleWriter;
use PhpOffice\PhpWord\Writer\Word2007\Style\Tab as TabStyleWriter;

/**
 * TOF element writer.
 *
 * @since 0.10.0
 */
class TOF extends AbstractElement
{
    /**
     * Write element.
     */
    public function write(): void
    {
        $xmlWriter = $this->getXmlWriter();
        $element = $this->getElement();
        if (!$element instanceof TOFElement) {
            return;
        }

        $captions = $element->getCaptions();
        $writeFieldMark = true;

        //////////////////////////////////////
        // Switch the comments of the below sections to have Word autoupdate the 
        // Table of Figures.
        //////////////////////////////////////
        // $xmlWriter->startElement('w:p');
        // $this->writeStyle($xmlWriter, $element);
        // $this->writeFieldMark($xmlWriter, $element);
        // $xmlWriter->endElement(); // w:p
        /////////////////////////////////////
        foreach ($captions as $caption) {
            $this->writeCaption($xmlWriter, $element, $caption, $writeFieldMark);
            if ($writeFieldMark) {
                $writeFieldMark = false;
            }
        }
        //////////////////////////////////////

        $xmlWriter->startElement('w:p');
        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'end');
        $xmlWriter->endElement();
        $xmlWriter->endElement();
        $xmlWriter->endElement();
    }

    /**
     * Write caption.
     */
    private function writeCaption(XMLWriter $xmlWriter, TOFElement $element, Caption $caption, bool $writeFieldMark): void
    {
        $tofStyle = $element->getStyleTOF();
        $fontStyle = $element->getStyleFont();
        $isObject = ($fontStyle instanceof Font) ? true : false;
        $paragraphStyle = $element->getParagraphStyle();
        $rId = $caption->getRelationId();
        $captionLabel = $caption->getLabel();
        $figureNumber = $caption->getFigureNumber();

        $xmlWriter->startElement('w:p');

        if (!empty($paragraphStyle)) {
            $this->writeParagraphStyle();
        }

        // Write style and field mark
        $this->writeStyle($xmlWriter, $element);
        if ($writeFieldMark) {
            $this->writeFieldMark($xmlWriter, $element);
        }

        // Hyperlink
        $xmlWriter->startElement('w:hyperlink');
        $xmlWriter->writeAttribute('w:anchor', "_{$captionLabel}{$rId}");
        $xmlWriter->writeAttribute('w:history', '1');

        // Title text
        $xmlWriter->startElement('w:r');
        if ($isObject) {
            $styleWriter = new FontStyleWriter($xmlWriter, $fontStyle);
            $styleWriter->write();
        }
        $xmlWriter->startElement('w:t');

        $captionText = $caption->getText();
        $this->writeText($captionLabel . ' ' . $figureNumber . '. ');
        $this->writeText(is_string($captionText) ? $captionText : '');

        $xmlWriter->endElement(); // w:t
        $xmlWriter->endElement(); // w:r

        $xmlWriter->startElement('w:r');
        $xmlWriter->writeElement('w:tab', null);
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'begin');
//        $xmlWriter->writeAttribute('w:dirty', 'true');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        // Commenting these lines out forces Word to regenerate the whole table.
        // Otherwise, neither the heading numbers nor the page numbers are displayed without
        // manually updating them.
        // $xmlWriter->startElement('w:r');
        // $xmlWriter->startElement('w:instrText');
        // $xmlWriter->writeAttribute('xml:space', 'preserve');
        // $xmlWriter->text(" PAGEREF _{$captionLabel}{$rId} \\h ");
        // $xmlWriter->endElement();
        // $xmlWriter->endElement();

        if ($caption->getPageNumber() !== null) {
            $xmlWriter->startElement('w:r');
            $xmlWriter->startElement('w:fldChar');
            $xmlWriter->writeAttribute('w:fldCharType', 'separate');
            $xmlWriter->endElement();
            $xmlWriter->endElement();

            $xmlWriter->startElement('w:r');
            $xmlWriter->startElement('w:t');
            $xmlWriter->text((string) $caption->getPageNumber());
            $xmlWriter->endElement();
            $xmlWriter->endElement();
        }

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'end');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->endElement(); // w:hyperlink

        $xmlWriter->endElement(); // w:p
    }

    /**
     * Write style.
     */
    private function writeStyle(XMLWriter $xmlWriter, TOFElement $element): void
    {
        $tofStyle = $element->getStyleTOF();
        $fontStyle = $element->getStyleFont();
        $isObject = ($fontStyle instanceof Font) ? true : false;

        $xmlWriter->startElement('w:pPr');

        // Paragraph
        if ($isObject && null !== $fontStyle->getParagraph()) {
            $styleWriter = new ParagraphStyleWriter($xmlWriter, $fontStyle->getParagraph());
            $styleWriter->write();
        }

        // Font
        if (!empty($fontStyle) && !$isObject) {
            $xmlWriter->startElement('w:rPr');
            $xmlWriter->startElement('w:rStyle');
            $xmlWriter->writeAttribute('w:val', $fontStyle);
            $xmlWriter->endElement();
            $xmlWriter->endElement(); // w:rPr
        }

        // Tab
        $xmlWriter->startElement('w:tabs');
        $styleWriter = new TabStyleWriter($xmlWriter, $tofStyle);
        $styleWriter->write();
        $xmlWriter->endElement();

        $xmlWriter->endElement(); // w:pPr
    }

    /**
     * Write TOF Field.
     */
    private function writeFieldMark(XMLWriter $xmlWriter, TOFElement $element): void
    {
        $TOFCaptionLabel = $element->getCaptionLabel();
        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'begin');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:instrText');
        $xmlWriter->writeAttribute('xml:space', 'preserve');
        $xmlWriter->writeRaw(' TOC \h \z \c "');
        $xmlWriter->text($TOFCaptionLabel);
        $xmlWriter->writeRaw('" ');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $xmlWriter->startElement('w:r');
        $xmlWriter->startElement('w:fldChar');
        $xmlWriter->writeAttribute('w:fldCharType', 'separate');
        $xmlWriter->endElement();
        $xmlWriter->endElement();
    }
}
