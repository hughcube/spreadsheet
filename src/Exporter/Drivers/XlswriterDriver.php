<?php

namespace HughCube\Spreadsheet\Exporter\Drivers;

use HughCube\Spreadsheet\Exporter\Coordinate;
use HughCube\Spreadsheet\Exporter\DriverInterface;
use HughCube\Spreadsheet\Exporter\Style;
use RuntimeException;

class XlswriterDriver implements DriverInterface
{
    /** @var \Vtiful\Kernel\Excel|null */
    protected $excel = null;

    protected bool $firstSheet = true;

    protected string $basename = '';

    /** @var array<string, resource> 样式资源缓存 */
    protected array $formatCache = [];

    /** @var array<int, array> 无样式行缓冲, 积攒后用 data() 批量写入 */
    protected array $rowBuffer = [];

    /** 缓冲区起始行号 */
    protected int $bufferStartRow = 0;

    /** 缓冲区大小阈值 */
    protected int $flushThreshold = 1000;

    public function open(string $filename): void
    {
        if (!extension_loaded('xlswriter')) {
            throw new RuntimeException('请安装 xlswriter 扩展: pecl install xlswriter');
        }

        $dir = dirname($filename);
        if (!is_dir($dir)) {
            mkdir($dir, 0755, true);
        }

        $config = ['path' => $dir . '/'];
        $this->excel = new \Vtiful\Kernel\Excel($config);
        $this->basename = basename($filename);
        $this->firstSheet = true;
        $this->formatCache = [];
        $this->rowBuffer = [];
        $this->bufferStartRow = 0;
    }

    public function addSheet(string $name): void
    {
        if ($this->firstSheet) {
            /** fileName() 第二参数设置第一个 sheet 名称 */
            $this->excel->fileName($this->basename, $name);
            $this->firstSheet = false;
        } else {
            $this->flushBuffer();
            $this->excel->addSheet($name);
        }
        $this->rowBuffer = [];
        $this->bufferStartRow = 0;
    }

    public function writeRow(int $rowIndex, array $values, array $cellStyles = []): void
    {
        if (empty($cellStyles)) {
            /** 无样式: 缓冲, 积攒后批量写入 */
            $this->rowBuffer[] = $values;
            if (count($this->rowBuffer) >= $this->flushThreshold) {
                $this->flushBuffer();
            }
        } else {
            /** 有样式: 先 flush 缓冲, 再逐单元格写入 */
            $this->flushBuffer();
            $this->writeStyledRow($rowIndex, $values, $cellStyles);
            /** 跳过此行, 更新缓冲起始位置 */
            $this->bufferStartRow = $rowIndex + 1;
        }
    }

    public function setColumnWidth(int $colIndex, float $width): void
    {
        $colLetter = Coordinate::indexToColumnLetter($colIndex);
        $range = "{$colLetter}:{$colLetter}";
        $this->excel->setColumn($range, $width);
    }

    public function setRowHeight(int $rowIndex, float $height): void
    {
        $this->flushBuffer();
        $cellAddr = 'A' . ($rowIndex + 1);
        $this->excel->setRow($cellAddr, $height);
    }

    public function mergeCells(int $startRow, int $startCol, int $endRow, int $endCol): void
    {
        $this->flushBuffer();
        $range = Coordinate::indexToColumnLetter($startCol) . ($startRow + 1)
            . ':' . Coordinate::indexToColumnLetter($endCol) . ($endRow + 1);
        $this->excel->mergeCells($range, '');
    }

    public function close(): void
    {
        $this->flushBuffer();
        $this->excel->output();
        $this->excel = null;
        $this->formatCache = [];
        $this->rowBuffer = [];
    }

    public function supportsStyle(): bool
    {
        return true;
    }

    /**
     * @return \Vtiful\Kernel\Excel|null
     */
    public function getNativeInstance(): mixed
    {
        return $this->excel;
    }

    /**
     * 将缓冲区中的无样式行用 data() 批量写入
     */
    protected function flushBuffer(): void
    {
        if (empty($this->rowBuffer)) {
            return;
        }

        /** 确保游标在正确位置 */
        $this->excel->setCurrentLine($this->bufferStartRow);

        /** data() 从游标位置批量写入, 一次调用写 N 行 */
        $this->excel->data($this->rowBuffer);

        $this->bufferStartRow += count($this->rowBuffer);
        $this->rowBuffer = [];
    }

    /**
     * 逐单元格写入带样式的行
     */
    protected function writeStyledRow(int $rowIndex, array $values, array $cellStyles): void
    {
        /** 检查是否所有单元格样式相同, 相同则只创建一次 format */
        $uniformFormat = null;
        $isUniform = true;
        $firstKey = null;
        foreach ($cellStyles as $colIdx => $style) {
            $key = $this->styleKey($style);
            if (null === $firstKey) {
                $firstKey = $key;
            } elseif ($key !== $firstKey) {
                $isUniform = false;
                break;
            }
        }

        if ($isUniform && null !== $firstKey && count($cellStyles) === count($values)) {
            /** 全行统一样式, 只取一次 format */
            $format = $this->getFormat($cellStyles[array_key_first($cellStyles)]);
            foreach ($values as $colIdx => $value) {
                $this->excel->insertText($rowIndex, $colIdx, $value, '', $format);
            }
        } else {
            /** 混合样式, 逐单元格处理 */
            foreach ($values as $colIdx => $value) {
                if (isset($cellStyles[$colIdx])) {
                    $format = $this->getFormat($cellStyles[$colIdx]);
                    $this->excel->insertText($rowIndex, $colIdx, $value, '', $format);
                } else {
                    $this->excel->insertText($rowIndex, $colIdx, $value);
                }
            }
        }
    }

    /**
     * 获取或创建 xlswriter Format 资源
     *
     * @return resource
     */
    protected function getFormat(Style $style)
    {
        $key = $this->styleKey($style);
        if (isset($this->formatCache[$key])) {
            return $this->formatCache[$key];
        }

        $handle = $this->excel->getHandle();
        $format = new \Vtiful\Kernel\Format($handle);

        if (true === $style->getBold()) {
            $format->bold();
        }
        if (true === $style->getItalic()) {
            $format->italic();
        }
        if (true === $style->getUnderline()) {
            $format->underline(\Vtiful\Kernel\Format::UNDERLINE_SINGLE);
        }
        if (true === $style->getStrikethrough()) {
            $format->strikeout();
        }
        if (null !== $style->getFontSize()) {
            $format->fontSize($style->getFontSize());
        }
        if (null !== $style->getFontFamily()) {
            $format->font($style->getFontFamily());
        }
        if (null !== $style->getFontColor()) {
            $format->fontColor($this->colorToInt($style->getFontColor()));
        }
        if (null !== $style->getBgColor()) {
            $format->background($this->colorToInt($style->getBgColor()));
        }
        if (null !== $style->getHorizontalAlign()) {
            $format->align($this->toXlswriterAlign($style->getHorizontalAlign()));
        }
        if (null !== $style->getVerticalAlign()) {
            $format->align($this->toXlswriterVerticalAlign($style->getVerticalAlign()));
        }
        if (true === $style->getWrapText()) {
            $format->wrap();
        }
        if (null !== $style->getNumberFormat()) {
            $format->number($style->getNumberFormat());
        }

        $resource = $format->toResource();
        $this->formatCache[$key] = $resource;
        return $resource;
    }

    protected function styleKey(Style $style): string
    {
        return md5(serialize($style->toArray()));
    }

    protected function colorToInt(string $hex): int
    {
        return (int) hexdec(ltrim($hex, '#'));
    }

    protected function toXlswriterAlign(string $align): int
    {
        return match ($align) {
            'left' => \Vtiful\Kernel\Format::FORMAT_ALIGN_LEFT,
            'center' => \Vtiful\Kernel\Format::FORMAT_ALIGN_CENTER,
            'right' => \Vtiful\Kernel\Format::FORMAT_ALIGN_RIGHT,
            default => \Vtiful\Kernel\Format::FORMAT_ALIGN_LEFT,
        };
    }

    protected function toXlswriterVerticalAlign(string $align): int
    {
        return match ($align) {
            'top' => \Vtiful\Kernel\Format::FORMAT_ALIGN_VERTICAL_TOP,
            'center' => \Vtiful\Kernel\Format::FORMAT_ALIGN_VERTICAL_CENTER,
            'bottom' => \Vtiful\Kernel\Format::FORMAT_ALIGN_VERTICAL_BOTTOM,
            default => \Vtiful\Kernel\Format::FORMAT_ALIGN_VERTICAL_CENTER,
        };
    }
}
