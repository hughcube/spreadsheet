<?php

namespace HughCube\Spreadsheet\Exporter\Drivers;

use HughCube\Spreadsheet\Exporter\DriverInterface;
use HughCube\Spreadsheet\Exporter\Style;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\CellAlignment;
use OpenSpout\Common\Entity\Style\CellVerticalAlignment;
use OpenSpout\Common\Entity\Style\Color;
use OpenSpout\Common\Entity\Style\Style as SpoutStyle;
use OpenSpout\Writer\XLSX\Options;
use OpenSpout\Writer\XLSX\Writer;
use RuntimeException;

class OpenSpoutDriver implements DriverInterface
{
    protected ?Writer $writer = null;

    protected ?Options $options = null;

    protected bool $firstSheet = true;

    protected int $sheetIndex = 0;

    /** @var float|null 当前行的预设行高 */
    protected ?float $pendingRowHeight = null;

    /** @var array<string, SpoutStyle> 样式缓存, 避免重复创建 immutable Style 对象 */
    protected array $styleCache = [];

    public function open(string $filename): void
    {
        if (!class_exists(Writer::class)) {
            throw new RuntimeException('请安装 openspout/openspout: composer require openspout/openspout');
        }

        $this->options = new Options();
        $this->writer = new Writer($this->options);
        $this->writer->openToFile($filename);
        $this->firstSheet = true;
        $this->sheetIndex = 0;
    }

    public function addSheet(string $name): void
    {
        if ($this->firstSheet) {
            $this->writer->getCurrentSheet()->setName($name);
            $this->firstSheet = false;
        } else {
            $this->sheetIndex++;
            $this->writer->addNewSheetAndMakeItCurrent()->setName($name);
        }
    }

    public function writeRow(int $rowIndex, array $values, array $cellStyles = []): void
    {
        if (empty($cellStyles)) {
            /** 无样式: 用 fromValues 一次性创建, 避免逐个构造 Cell 对象 */
            $row = Row::fromValues($values);
        } else {
            $cells = [];
            foreach ($values as $colIdx => $value) {
                if (isset($cellStyles[$colIdx])) {
                    $cells[] = Cell::fromValue($value, $this->convertStyle($cellStyles[$colIdx]));
                } else {
                    $cells[] = Cell::fromValue($value);
                }
            }
            $row = new Row($cells);
        }

        if (null !== $this->pendingRowHeight) {
            $row = $row->withHeight($this->pendingRowHeight);
            $this->pendingRowHeight = null;
        }

        $this->writer->addRow($row);
    }

    public function setColumnWidth(int $colIndex, float $width): void
    {
        /** OpenSpout Options 列索引 1-based */
        $this->options->setColumnWidth($width, $colIndex + 1);
    }

    public function setRowHeight(int $rowIndex, float $height): void
    {
        /** 缓存行高, 在下次 writeRow 时应用到 Row 对象 */
        $this->pendingRowHeight = $height;
    }

    public function mergeCells(int $startRow, int $startCol, int $endRow, int $endCol): void
    {
        /** OpenSpout Options mergeCells 使用 1-based 索引 */
        $this->options->mergeCells($startCol + 1, $startRow + 1, $endCol + 1, $endRow + 1, $this->sheetIndex);
    }

    public function close(): void
    {
        $this->writer->close();
        $this->writer = null;
        $this->options = null;
        $this->styleCache = [];
    }

    public function supportsStyle(): bool
    {
        return true;
    }

    /**
     * @return Writer|null
     */
    public function getNativeInstance(): mixed
    {
        return $this->writer;
    }

    protected function convertStyle(Style $style): SpoutStyle
    {
        $key = md5(serialize($style->toArray()));
        if (isset($this->styleCache[$key])) {
            return $this->styleCache[$key];
        }

        $spoutStyle = new SpoutStyle();

        if (null !== $style->getBold()) {
            $spoutStyle = $spoutStyle->withFontBold($style->getBold());
        }
        if (null !== $style->getItalic()) {
            $spoutStyle = $spoutStyle->withFontItalic($style->getItalic());
        }
        if (null !== $style->getUnderline()) {
            $spoutStyle = $spoutStyle->withFontUnderline($style->getUnderline());
        }
        if (null !== $style->getStrikethrough()) {
            $spoutStyle = $spoutStyle->withFontStrikethrough($style->getStrikethrough());
        }
        if (null !== $style->getFontSize()) {
            $spoutStyle = $spoutStyle->withFontSize((int) $style->getFontSize());
        }
        if (null !== $style->getFontFamily()) {
            $spoutStyle = $spoutStyle->withFontName($style->getFontFamily());
        }
        if (null !== $style->getFontColor()) {
            $spoutStyle = $spoutStyle->withFontColor($this->hexToColor($style->getFontColor()));
        }
        if (null !== $style->getBgColor()) {
            $spoutStyle = $spoutStyle->withBackgroundColor($this->hexToColor($style->getBgColor()));
        }
        if (null !== $style->getWrapText()) {
            $spoutStyle = $spoutStyle->withShouldWrapText($style->getWrapText());
        }
        if (null !== $style->getHorizontalAlign()) {
            $spoutStyle = $spoutStyle->withCellAlignment($this->toAlignment($style->getHorizontalAlign()));
        }
        if (null !== $style->getVerticalAlign()) {
            $vAlign = $this->toVerticalAlignment($style->getVerticalAlign());
            $spoutStyle = $spoutStyle->withCellVerticalAlignment($vAlign);
        }
        if (null !== $style->getNumberFormat()) {
            $spoutStyle = $spoutStyle->withFormat($style->getNumberFormat());
        }

        $this->styleCache[$key] = $spoutStyle;
        return $spoutStyle;
    }

    protected function hexToColor(string $hex): string
    {
        $hex = ltrim($hex, '#');
        $r = (int) hexdec(substr($hex, 0, 2));
        $g = (int) hexdec(substr($hex, 2, 2));
        $b = (int) hexdec(substr($hex, 4, 2));
        return Color::rgb($r, $g, $b);
    }

    protected function toAlignment(string $align): CellAlignment
    {
        return match ($align) {
            'left' => CellAlignment::LEFT,
            'center' => CellAlignment::CENTER,
            'right' => CellAlignment::RIGHT,
            'justify' => CellAlignment::JUSTIFY,
            default => CellAlignment::LEFT,
        };
    }

    protected function toVerticalAlignment(string $align): CellVerticalAlignment
    {
        return match ($align) {
            'top' => CellVerticalAlignment::TOP,
            'center' => CellVerticalAlignment::CENTER,
            'bottom' => CellVerticalAlignment::BOTTOM,
            default => CellVerticalAlignment::CENTER,
        };
    }
}
