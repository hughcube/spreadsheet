<?php

namespace HughCube\Spreadsheet\Exporter\Drivers;

use HughCube\Spreadsheet\Exporter\DriverInterface;
use HughCube\Spreadsheet\Exporter\Style;
use RuntimeException;

class CsvDriver implements DriverInterface
{
    /** @var resource|null */
    protected $handle = null;

    protected string $separator;

    protected string $enclosure;

    public function __construct(string $separator = ',', string $enclosure = '"')
    {
        $this->separator = $separator;
        $this->enclosure = $enclosure;
    }

    public function open(string $filename): void
    {
        $this->handle = fopen($filename, 'wb');
        if (false === $this->handle) {
            throw new RuntimeException("无法打开文件: {$filename}");
        }

        /** UTF-8 BOM, 保证 Excel 正确识别编码 */
        fwrite($this->handle, "\xEF\xBB\xBF");
    }

    public function addSheet(string $name): void
    {
        /** CSV 不支持多 sheet */
    }

    public function writeRow(int $rowIndex, array $values, array $cellStyles = []): void
    {
        fputcsv($this->handle, $values, $this->separator, $this->enclosure);
    }

    public function setColumnWidth(int $colIndex, float $width): void
    {
    }

    public function setRowHeight(int $rowIndex, float $height): void
    {
    }

    public function mergeCells(int $startRow, int $startCol, int $endRow, int $endCol): void
    {
    }

    public function close(): void
    {
        if (null !== $this->handle) {
            fclose($this->handle);
            $this->handle = null;
        }
    }

    public function supportsStyle(): bool
    {
        return false;
    }

    /**
     * @return resource|null
     */
    public function getNativeInstance(): mixed
    {
        return $this->handle;
    }
}
