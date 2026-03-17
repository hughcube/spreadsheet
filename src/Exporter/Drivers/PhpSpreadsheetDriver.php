<?php

namespace HughCube\Spreadsheet\Exporter\Drivers;

use HughCube\Spreadsheet\Exporter\Coordinate;
use HughCube\Spreadsheet\Exporter\DriverInterface;
use HughCube\Spreadsheet\Exporter\Style;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use PhpOffice\PhpSpreadsheet\Writer\Xls as XlsWriter;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use RuntimeException;

class PhpSpreadsheetDriver implements DriverInterface
{
    protected ?Spreadsheet $spreadsheet = null;

    protected string $filename = '';

    protected int $sheetIndex = -1;

    public function open(string $filename): void
    {
        if (!class_exists(Spreadsheet::class)) {
            throw new RuntimeException('请安装 phpoffice/phpspreadsheet: composer require phpoffice/phpspreadsheet');
        }

        $this->spreadsheet = new Spreadsheet();
        $this->spreadsheet->removeSheetByIndex(0);
        $this->filename = $filename;
    }

    public function addSheet(string $name): void
    {
        $this->sheetIndex++;
        $sheet = new Worksheet($this->spreadsheet, $name);
        $this->spreadsheet->addSheet($sheet, $this->sheetIndex);
        $this->spreadsheet->setActiveSheetIndex($this->sheetIndex);
    }

    public function writeRow(int $rowIndex, array $values, array $cellStyles = []): void
    {
        $sheet = $this->spreadsheet->getActiveSheet();
        $excelRow = $rowIndex + 1;

        foreach ($values as $colIdx => $value) {
            $colLetter = Coordinate::indexToColumnLetter($colIdx);
            $cellAddr = $colLetter . $excelRow;
            $sheet->setCellValue($cellAddr, $value);

            if (isset($cellStyles[$colIdx])) {
                $this->applyStyle($sheet->getStyle($cellAddr), $cellStyles[$colIdx]);
            }
        }
    }

    public function setColumnWidth(int $colIndex, float $width): void
    {
        $colLetter = Coordinate::indexToColumnLetter($colIndex);
        $this->spreadsheet->getActiveSheet()
            ->getColumnDimension($colLetter)
            ->setWidth($width);
    }

    public function setRowHeight(int $rowIndex, float $height): void
    {
        $this->spreadsheet->getActiveSheet()
            ->getRowDimension($rowIndex + 1)
            ->setRowHeight($height);
    }

    public function mergeCells(int $startRow, int $startCol, int $endRow, int $endCol): void
    {
        $range = Coordinate::indexToColumnLetter($startCol) . ($startRow + 1)
            . ':' . Coordinate::indexToColumnLetter($endCol) . ($endRow + 1);
        $this->spreadsheet->getActiveSheet()->mergeCells($range);
    }

    public function close(): void
    {
        $ext = strtolower(pathinfo($this->filename, PATHINFO_EXTENSION));
        $writer = match ($ext) {
            'xls' => new XlsWriter($this->spreadsheet),
            'csv' => new CsvWriter($this->spreadsheet),
            default => new XlsxWriter($this->spreadsheet),
        };

        $writer->save($this->filename);
        $this->spreadsheet->disconnectWorksheets();
        $this->spreadsheet = null;
    }

    public function supportsStyle(): bool
    {
        return true;
    }

    /**
     * @return Spreadsheet|null
     */
    public function getNativeInstance(): mixed
    {
        return $this->spreadsheet;
    }

    protected function applyStyle(\PhpOffice\PhpSpreadsheet\Style\Style $target, Style $style): void
    {
        $font = $target->getFont();

        if (null !== $style->getBold()) {
            $font->setBold($style->getBold());
        }
        if (null !== $style->getItalic()) {
            $font->setItalic($style->getItalic());
        }
        if (null !== $style->getUnderline()) {
            $font->setUnderline(
                $style->getUnderline()
                    ? \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLE
                    : \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE
            );
        }
        if (null !== $style->getStrikethrough()) {
            $font->setStrikethrough($style->getStrikethrough());
        }
        if (null !== $style->getFontSize()) {
            $font->setSize($style->getFontSize());
        }
        if (null !== $style->getFontFamily()) {
            $font->setName($style->getFontFamily());
        }
        if (null !== $style->getFontColor()) {
            $font->getColor()->setRGB(ltrim($style->getFontColor(), '#'));
        }

        if (null !== $style->getBgColor()) {
            $target->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB(ltrim($style->getBgColor(), '#'));
        }

        $alignment = $target->getAlignment();
        if (null !== $style->getHorizontalAlign()) {
            $alignment->setHorizontal(match ($style->getHorizontalAlign()) {
                'left' => Alignment::HORIZONTAL_LEFT,
                'center' => Alignment::HORIZONTAL_CENTER,
                'right' => Alignment::HORIZONTAL_RIGHT,
                default => $style->getHorizontalAlign(),
            });
        }
        if (null !== $style->getVerticalAlign()) {
            $alignment->setVertical(match ($style->getVerticalAlign()) {
                'top' => Alignment::VERTICAL_TOP,
                'center' => Alignment::VERTICAL_CENTER,
                'bottom' => Alignment::VERTICAL_BOTTOM,
                default => $style->getVerticalAlign(),
            });
        }
        if (null !== $style->getWrapText()) {
            $alignment->setWrapText($style->getWrapText());
        }

        if (null !== $style->getNumberFormat()) {
            $target->getNumberFormat()->setFormatCode($style->getNumberFormat());
        }
    }
}
