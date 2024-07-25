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
use PhpOffice\PhpWord\Style\TOC as TOCStyle;

/**
 * Table of contents.
 */
class TOC extends AbstractElement
{
    /**
     * TOC style.
     *
     * @var \PhpOffice\PhpWord\Style\TOC
     */
    private $tocStyle;

    /**
     * Font style.
     *
     * @var \PhpOffice\PhpWord\Style\Font|string
     */
    private $fontStyle;

    /**
     * Min title depth to show.
     *
     * @var int
     */
    private $minDepth = 1;

    /**
     * Max title depth to show.
     *
     * @var int
     */
    private $maxDepth = 9;

    /**
     * t field heading pair string.
     *
     * @var string
     */
    private $tField = '';

    /**
     * t field heading pair array.
     *
     * @var array
     */
    private $tFieldArray = [];

    /**
     * String containing TOC switches.
     *
     * @var string
     */
    private $switchString = '\u';

    /**
     * Create a new Table-of-Contents Element.
     *
     * @param mixed $fontStyle
     * @param array $tocStyle
     * @param int $minDepth
     * @param int $maxDepth
     * @param string $switchString
     * @param string $tField
     */
    public function __construct($fontStyle = null, $tocStyle = null, $minDepth = 1, $maxDepth = 9, $switchString = '\h \z \u', $tField = '')
    {
        $this->tocStyle = new TOCStyle();

        if (null !== $tocStyle && is_array($tocStyle)) {
            $this->tocStyle->setStyleByArray($tocStyle);
        }

        if (null !== $fontStyle && is_array($fontStyle)) {
            $this->fontStyle = new Font();
            $this->fontStyle->setStyleByArray($fontStyle);
        } else {
            $this->fontStyle = $fontStyle;
        }

        $this->minDepth = $minDepth;
        $this->maxDepth = $maxDepth;

        $this->switchString = $switchString;

        $this->setTField($tField);
    }

    /**
     * Get all titles.
     *
     * @return array
     */
    public function getTitles()
    {
        if (!$this->phpWord instanceof PhpWord) {
            return [];
        }

        $titles = $this->phpWord->getTitles()->getItems();
        foreach ($titles as $i => $title) {
            /** @var \PhpOffice\PhpWord\Element\Title $title Type hint */
            $styleName = $title->getStyle();
            $depth = $title->getDepth();

            if (array_key_exists($styleName, $this->tFieldArray)) {
                $title->setDepth($this->tFieldArray[$styleName]);
            } elseif (!empty($styleName) && array_key_exists(str_replace('Heading', 'Heading ', $styleName), $this->tFieldArray)) {
                $title->setDepth($this->tFieldArray[str_replace('Heading', 'Heading ', $styleName)]);
            } else {
                if ($this->minDepth > $depth) {
                    unset($titles[$i]);
                }
                if (($this->maxDepth != 0) && ($this->maxDepth < $depth)) {
                    unset($titles[$i]);
                }
            }
        }

        return $titles;
    }

    /**
     * Get TOC Style.
     *
     * @return \PhpOffice\PhpWord\Style\TOC
     */
    public function getStyleTOC()
    {
        return $this->tocStyle;
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
     * Set max depth.
     *
     * @param int $value
     */
    public function setMaxDepth($value): void
    {
        $this->maxDepth = $value;
    }

    /**
     * Get Max Depth.
     *
     * @return int Max depth of titles
     */
    public function getMaxDepth()
    {
        return $this->maxDepth;
    }

    /**
     * Set min depth.
     *
     * @param int $value
     */
    public function setMinDepth($value): void
    {
        $this->minDepth = $value;
    }

    /**
     * Get Min Depth.
     *
     * @return int Min depth of titles
     */
    public function getMinDepth()
    {
        return $this->minDepth;
    }

    /**
     * Get switchString.
     *
     * @return string switchString
     */
    public function getSwitchString()
    {
        return $this->switchString;
    }

    /**
     * Get t field string.
     *
     * @return string tField
     */
    public function getTField()
    {
        return $this->tField;
    }

    /**
     * Get t field array.
     *
     * @return array tFieldArray
     */
    public function getTFieldArray()
    {
        return $this->tFieldArray;
    }

    /**
     * Set t field array.
     *
     * @param string $tField
     *
     * @return self
     */
    public function setTField($tField = '')
    {
        $this->tField = $tField;
        $this->tFieldArray = [];
        if (!empty($tField)) {
            $arr = explode(',', $tField);
            for ($i = 0; $i < count($arr); $i = $i + 2) {
                $this->tFieldArray[$arr[$i]] = $arr[$i + 1];
            }
        }

        return $this;
    }
}
