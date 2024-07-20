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

namespace PhpOffice\PhpWord\Style;

/**
 * TOF style.
 */
class TOF extends Tab
{
    /**
     * Create a new TOF Style.
     */
    public function __construct()
    {
        parent::__construct(self::TAB_STOP_RIGHT, 9360, self::TAB_LEADER_DOT);
    }

    /**
     * Get Tab Position.
     *
     * @return float|int
     */
    public function getTabPos()
    {
        return $this->getPosition();
    }

    /**
     * Set Tab Position.
     *
     * @param float|int $value
     *
     * @return \PhpOffice\PhpWord\Style\Tab
     */
    public function setTabPos($value)
    {
        return $this->setPosition($value);
    }

    /**
     * Get Tab Leader.
     *
     * @return string
     */
    public function getTabLeader()
    {
        return $this->getLeader();
    }

    /**
     * Set Tab Leader.
     *
     * @param string $value
     *
     * @return \PhpOffice\PhpWord\Style\Tab
     */
    public function setTabLeader($value = self::TAB_LEADER_DOT)
    {
        return $this->setLeader($value);
    }
}
