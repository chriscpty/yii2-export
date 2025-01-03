<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2023
 * @package yii2-export
 * @version 1.4.3
 */

namespace kartik\export;

use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\CellVerticalAlignment;
use OpenSpout\Common\Entity\Style\Color;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

/**
 * This class contains helper methods for using PhpSpreadsheet data with Openspout.
 */
class OpenspoutHelper
{
    /**
     * Take an options array as used by `PhpSpreadsheet`'s `applyFromArray` and build an according Openspout `Style` object.
     * @param array $opts The options, as formatted for `PhpSpreadsheet` `applyFromArray`.
     * @return Style The resulting style.
     * @throws InvalidArgumentException
     * @see \PhpOffice\PhpSpreadsheet\Style\Style::applyFromArray() for available options.
     * TODO - this doesn't fully support all options PhpSpreadsheet does.
     * TODO - figure out specifics of $opts['font']['color']
     */
    public static function createStyleFromPhpSpreadsheetOptions(array $opts): Style {
        $style = new Style();
        if (isset($opts['font'])) {
            if (!empty($opts['font']['name'])) {
                $style->setFontName($opts['font']['name']);
            }
            if (!empty($opts['font']['bold'])) {
                $style->setFontBold();
            }
            if (!empty($opts['font']['italic'])) {
                $style->setFontItalic();
            }
            if (!empty($opts['font']['underline'])) {
                $style->setFontUnderline();
            }
            if (!empty($opts['font']['strikethrough'])) {
                $style->setFontStrikethrough();
            }
            if (!empty($opts['font']['color']['rgb'])) {
                $color = $opts['font']['color']['rgb'];
                $style->setFontColor(Color::rgb((int)substr($color, 0, 2), (int)substr($color, 2, 2), (int)substr($color, 4, 2)));
            }
        }
        if (isset($opts['alignment'])) {
            if (isset($opts['alignment']['horizontal'])) {
                $style->setCellAlignment(self::mapPhpSpreadsheetHorizontalAlignmentConstant($opts['alignment']['horizontal']));
            }
            if (isset($opts['alignment']['vertical'])) {
                $style->setCellVerticalAlignment(self::mapPhpSpreadsheetVerticalAlignmentConstant($opts['alignment']['vertical']));
            }
            if (!empty($opts['alignment']['wrapText'])) {
                $style->setShouldWrapText(true);
            }
        }
        return $style;
    }

    /**
     * Map PhpSpreadsheet constants to Openspout ones.
     * Note that because Openspout will throw an error for alignments other than "Left", "Right", "Justify" and "Center",
     * other values are mapped to "Center" rather than a proper equivalent.
     * @param string $constant
     * @return string
     */
    public static function mapPhpSpreadsheetHorizontalAlignmentConstant(string $constant): string
    {
        return match ($constant) {
            Alignment::HORIZONTAL_LEFT => CellAlignment::LEFT,
            Alignment::HORIZONTAL_RIGHT => CellAlignment::RIGHT,
            Alignment::HORIZONTAL_JUSTIFY => CellAlignment::JUSTIFY,
            default => CellAlignment::CENTER,
        };
    }

    /**
     * Map PhpSpreadsheet constants to Openspout ones.
     * @param string $constant
     * @return string
     */
    public static function mapPhpSpreadsheetVerticalAlignmentConstant(string $constant): string
    {
        return match ($constant) {
            Alignment::VERTICAL_TOP => CellVerticalAlignment::TOP,
            Alignment::VERTICAL_BOTTOM => CellVerticalAlignment::BOTTOM,
            Alignment::VERTICAL_DISTRIBUTED => CellVerticalAlignment::DISTRIBUTED,
            Alignment::VERTICAL_JUSTIFY => CellVerticalAlignment::JUSTIFY,
            default => CellVerticalAlignment::CENTER,
        };
    }
}