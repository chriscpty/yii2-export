<?php

/**
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2015 - 2023
 * @package yii2-export
 * @version 1.4.3
 */

namespace kartik\export;

use OpenSpout\Common\Entity\Style\Border;
use OpenSpout\Common\Entity\Style\BorderPart;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\CellVerticalAlignment;
use OpenSpout\Common\Entity\Style\Color;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\InvalidArgumentException;
use OpenSpout\Writer\Exception\Border\InvalidNameException;
use OpenSpout\Writer\Exception\Border\InvalidStyleException;
use OpenSpout\Writer\Exception\Border\InvalidWidthException;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border as PhpSpreadsheetBorder;

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
     * @throws InvalidNameException
     * @throws InvalidStyleException
     * @throws InvalidWidthException
     * @see \PhpOffice\PhpSpreadsheet\Style\Style::applyFromArray() for available options.
     * TODO - this doesn't fully support all options PhpSpreadsheet does.
     * TODO - figure out specifics of $opts['font']['color']
     */
    public static function createStyleFromPhpSpreadsheetOptions(array $opts): Style
    {
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
            if (!empty($opts['font']['color'])) {
                $color = $opts['font']['color'];
                $style->setFontColor($color['rgb'] ?? $color['argb'] ?? $color);
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
        if (!empty($opts['borders'])) {
            $borderParts = [];
            foreach ([Border::TOP, Border::LEFT, Border::BOTTOM, Border::RIGHT] as $name) {
                if (isset($opts['borders'][$name])) {
                    $color = $opts['borders'][$name]['color']['rgb'] ?? Color::BLACK;
                    $width = Border::WIDTH_MEDIUM;
                    $style = Border::STYLE_NONE;
                    if (isset($opts['borders'][$name]['borderStyle'])) {
                        [$width, $style] = self::mapPhpSpreadsheetBorderConstant($opts['borders'][$name]['borderStyle']);
                    }
                    $borderParts[] = new BorderPart($name, $color, $width, $style);
                }
            }
            $style->setBorder(new Border(...$borderParts));
        }
        if (isset($opts['fill']['fillType']) && $opts['fill']['fillType'] === 'solid') {
            $color = $opts['fill']['color'];
            $style->setBackgroundColor($color['rgb'] ?? $color['argb'] ?? $color);
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
    protected static function mapPhpSpreadsheetHorizontalAlignmentConstant(string $constant): string
    {
        return match ($constant) {
            Alignment::HORIZONTAL_LEFT    => CellAlignment::LEFT,
            Alignment::HORIZONTAL_RIGHT   => CellAlignment::RIGHT,
            Alignment::HORIZONTAL_JUSTIFY => CellAlignment::JUSTIFY,
            default                       => CellAlignment::CENTER,
        };
    }

    /**
     * Map PhpSpreadsheet constants to Openspout ones.
     * @param string $constant
     * @return string
     */
    protected static function mapPhpSpreadsheetVerticalAlignmentConstant(string $constant): string
    {
        return match ($constant) {
            Alignment::VERTICAL_TOP         => CellVerticalAlignment::TOP,
            Alignment::VERTICAL_BOTTOM      => CellVerticalAlignment::BOTTOM,
            Alignment::VERTICAL_DISTRIBUTED => CellVerticalAlignment::DISTRIBUTED,
            Alignment::VERTICAL_JUSTIFY     => CellVerticalAlignment::JUSTIFY,
            default                         => CellVerticalAlignment::CENTER,
        };
    }

    /**
     * Map PhpSpreadsheet constants to Openspout ones.
     * Note that Openspout's supported border types are limited by comparison, so all "Dashdot" styles are just turned into medium dashed borders.
     * @param string $constant
     * @return array
     */
    protected static function mapPhpSpreadsheetBorderConstant(string $constant): array
    {
        return match ($constant) {
            PhpSpreadsheetBorder::BORDER_NONE                                    => [Border::WIDTH_MEDIUM, Border::STYLE_NONE],
            PhpSpreadsheetBorder::BORDER_DOTTED                                  => [Border::WIDTH_MEDIUM, Border::STYLE_DOTTED],
            PhpSpreadsheetBorder::BORDER_DOUBLE                                  => [Border::WIDTH_MEDIUM, Border::STYLE_DOUBLE],
            PhpSpreadsheetBorder::BORDER_HAIR, PhpSpreadsheetBorder::BORDER_THIN => [Border::WIDTH_THIN, Border::STYLE_SOLID],
            PhpSpreadsheetBorder::BORDER_MEDIUM                                  => [Border::WIDTH_MEDIUM, Border::STYLE_SOLID],
            PhpSpreadsheetBorder::BORDER_SLANTDASHDOT                            => [Border::WIDTH_THIN, Border::STYLE_DASHED],
            PhpSpreadsheetBorder::BORDER_THICK                                   => [Border::WIDTH_THICK, Border::STYLE_SOLID],
            PhpSpreadsheetBorder::BORDER_OMIT                                    => [Border::WIDTH_THIN, Border::STYLE_NONE],
            default                                                              => [Border::WIDTH_MEDIUM, Border::STYLE_DASHED],
        };
    }
}