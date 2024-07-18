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

namespace PhpOffice\PhpWord\Element;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Style\Paragraph;
use PhpOffice\PhpWord\Style\TOF as TOFStyle;

/**
 * Table of figures.
 */
class TOF extends AbstractElement
{
    /**
     * Caption label.
     *
     * @var string
     */
    private $captionLabel;

    /**
     * TOF style.
     *
     * @var \PhpOffice\PhpWord\Style\TOF
     */
    private $tofStyle;

    /**
     * Font style.
     *
     * @var \PhpOffice\PhpWord\Style\Font|string
     */
    private $fontStyle;

    /**
     * Paragraph style.
     *
     * @var \PhpOffice\PhpWord\Style\Paragraph|string
     */
    private $paragraphStyle;

    /**
     * Create a new Table-of-Figures Element.
     *
     * @param string $captionLabel
     * @param array $tofStyle
     * @param mixed $fontStyle
     * @param mixed $paragraphStyle
     */
    public function __construct($captionLabel = 'Figure', $tofStyle = null, $fontStyle = null, $paragraphStyle = null)
    {
        $this->captionLabel = $captionLabel;

        $this->tofStyle = new TOFStyle();

        if (null !== $tofStyle && is_array($tofStyle)) {
            $this->tofStyle->setStyleByArray($tofStyle);
        }

        if (null !== $fontStyle && is_array($fontStyle)) {
            $this->fontStyle = new Font();
            $this->fontStyle->setStyleByArray($fontStyle);
        } else {
            $this->fontStyle = $fontStyle;
        }

        $paragraphStyle = $this->setParagraphStyle($paragraphStyle);
        $this->setFontStyle($fontStyle, $paragraphStyle);
    }

    /**
     * Get all captions.
     *
     * @return array
     */
    public function getCaptions()
    {
        if (!$this->phpWord instanceof PhpWord) {
            return [];
        }

        $captions = $this->phpWord->getCaptions()->getItems();
        foreach ($captions as $i => $caption) {
            /** @var \PhpOffice\PhpWord\Element\Caption $caption Type hint */
            $label = $caption->getLabel();
            if ($this->captionLabel !== $label) {
                unset($captions[$i]);
            }
        }

        return $captions;
    }

    /**
     * Get TOF Style.
     *
     * @return \PhpOffice\PhpWord\Style\TOF
     */
    public function getStyleTOF()
    {
        return $this->tofStyle;
    }

    /**
     * Get Font Style.
     *
     * @return \PhpOffice\PhpWord\Style\Font|string
     */
    public function getStyleFont()
    {
        return $this->fontStyle;
    }

    /**
     * Get Caption Label.
     *
     * @return int Max depth of titles
     */
    public function getCaptionLabel()
    {
        return $this->captionLabel;
    }

    /**
     * Set Text style.
     *
     * @param array|\PhpOffice\PhpWord\Style\Font|string $style
     * @param array|\PhpOffice\PhpWord\Style\Paragraph|string $paragraphStyle
     *
     * @return \PhpOffice\PhpWord\Style\Font|string
     */
    public function setFontStyle($style = null, $paragraphStyle = null)
    {
        if ($style instanceof Font) {
            $this->fontStyle = $style;
            $this->setParagraphStyle($paragraphStyle);
        } elseif (is_array($style)) {
            $this->fontStyle = new Font('text', $paragraphStyle);
            $this->fontStyle->setStyleByArray($style);
        } elseif (null === $style) {
            $this->fontStyle = new Font('text', $paragraphStyle);
        } else {
            $this->fontStyle = $style;
            $this->setParagraphStyle($paragraphStyle);
        }

        return $this->fontStyle;
    }

    /**
     * Get Text style.
     *
     * @return \PhpOffice\PhpWord\Style\Font|string
     */
    public function getFontStyle()
    {
        return $this->fontStyle;
    }

    /**
     * Get Paragraph style.
     *
     * @return \PhpOffice\PhpWord\Style\Paragraph|string
     */
    public function getParagraphStyle()
    {
        return $this->paragraphStyle;
    }

    /**
     * Set Paragraph style.
     *
     * @param array|\PhpOffice\PhpWord\Style\Paragraph|string $style
     *
     * @return \PhpOffice\PhpWord\Style\Paragraph|string
     */
    public function setParagraphStyle($style = null)
    {
        if (is_array($style)) {
            $this->paragraphStyle = new Paragraph();
            $this->paragraphStyle->setStyleByArray($style);
        } elseif ($style instanceof Paragraph) {
            $this->paragraphStyle = $style;
        } elseif (null === $style) {
            $this->paragraphStyle = new Paragraph();
        } else {
            $this->paragraphStyle = $style;
        }

        return $this->paragraphStyle;
    }
}
