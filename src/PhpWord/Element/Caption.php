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

/**
 * Caption element.
 */
class Caption extends Text
{
    /**
     * figureNumbers.
     *
     * @var array
     */
    protected static $figureNumbers = [];

    /**
     * Label.
     *
     * @var string
     */
    protected $label;

    /**
     * figureNumber.
     *
     * @var string
     */
    protected $figureNumber;

    /**
     * Page number.
     *
     * @var int
     */
    private $pageNumber;

    /**
     * Is part of collection.
     *
     * @var bool
     */
    protected $collectionRelation = true;

    /**
     * Create a new Caption Element.
     *
     * @param string $text
     * @param mixed $label (or change to $captionStyle[] in the future)
     * @param mixed $fontStyle
     * @param mixed $paragraphStyle
     */
    public function __construct($label = 'Figure', $text = null, $fontStyle = null, $paragraphStyle = null, ?int $pageNumber = null)
    {
        $this->setLabel($label);
        $this->setText($text);
        $paragraphStyle = $this->setParagraphStyle($paragraphStyle);
        $this->setFontStyle($fontStyle, $paragraphStyle);
        $this->setFigureNumber();

        if ($pageNumber !== null) {
            $this->pageNumber = $pageNumber;
        }
    }

    /**
     * Set Caption label.
     *
     * @param string $label
     *
     * @return self
     */
    public function setLabel($label)
    {
        $this->label = $label;

        return $this;
    }

    /**
     * Get Caption label.
     */
    public function getLabel(): string
    {
        return $this->label;
    }

    /**
     * Get page number.
     */
    public function getPageNumber(): ?int
    {
        return $this->pageNumber;
    }

    /**
     * Get figure number.
     */
    public function getFigureNumber(): string
    {
        return $this->figureNumber;
    }

    /**
     * Set figure number.
     */
    public function setFigureNumber(): void
    {
        if (array_key_exists($this->label, self::$figureNumbers)) {
            ++self::$figureNumbers[$this->label];
        } else {
            self::$figureNumbers[$this->label] = 1;
        }
        $this->figureNumber = (string) self::$figureNumbers[$this->label];
    }
}
