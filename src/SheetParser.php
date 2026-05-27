<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/5
 * Time: 23:30
 */

namespace HughCube\Spreadsheet;

use HughCube\Spreadsheet\Models\Headers;
use PhpOffice\PhpSpreadsheet\Cell\CellAddress;
use PhpOffice\PhpSpreadsheet\Cell\CellRange;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Throwable;

class SheetParser
{
    /**
     * @var Worksheet
     */
    protected $sheet;

    /**
     * @var false|null|Headers
     */
    protected $headers = false;

    /**
     * @var null|Headers
     */
    protected $closestHeaders = null;

    /**
     * @var array
     */
    protected $headerPatterns;

    /**
     * @var array
     */
    protected $errors = [];

    /**
     * @return static
     */
    public static function parse(Worksheet $sheet, array $patterns = []): SheetParser
    {
        /** @phpstan-ignore-next-line */
        return new static($sheet, $patterns);
    }

    protected function __construct(Worksheet $sheet, array $patterns)
    {
        $this->sheet = $sheet;
        $this->headerPatterns = $patterns;
    }

    public function getSheet(): Worksheet
    {
        return $this->sheet;
    }

    public function getHeaderPatterns(): array
    {
        return $this->headerPatterns;
    }

    /**
     * @param  int|string  $name
     * @param  null|callable  $is
     * @param  null|callable  $format
     * @param  bool  $required
     * @return $this
     */
    public function addHeaderPattern($name, $is, $format = null, bool $required = true): SheetParser
    {
        $this->headerPatterns[$name] = ['is' => $is, 'format' => $format, 'required' => $required];

        return $this;
    }

    /**
     * @return null|Headers
     */
    public function getHeaders()
    {
        if (false === $this->headers) {
            list($this->headers, $this->closestHeaders) = Headers::parse($this, $this->getHeaderPatterns());
        }
        return $this->headers;
    }

    public function getClosestHeaders(): ?Headers
    {
        $this->getHeaders();

        return $this->closestHeaders;
    }

    public function getErrors(): array
    {
        return $this->errors;
    }

    public function getDataIterator($maxRowCount = null)
    {
        $headers = $this->getHeaders();
        if (null == $headers) {
            return;
        }
        $endColumn = $headers->getMaxHeaderIndex();

        $startRowIndex = $headers->getIndex() + 1;
        $endRowIndex = empty($maxRowCount) ? null : $startRowIndex + $maxRowCount;
        foreach ($this->getsheet()->getRowIterator($startRowIndex, $endRowIndex) as $row) {
            $cells = [];
            foreach ($row->getCellIterator('A', $endColumn) as $index => $cell) {
                try {
                    /** 数字/公式且非日期 → 原始/计算值, 避免千分位等数字格式被 getFormattedValue 转成 "1,000" 这类字符串, 调用方 floatval 时会截断成 1 */
                    $dataType = $cell->getDataType();
                    $isNumericLike = $dataType === DataType::TYPE_NUMERIC || $dataType === DataType::TYPE_FORMULA;
                    if ($isNumericLike && !Date::isDateTime($cell)) {
                        $cells[$index] = $cell->getCalculatedValue();
                    } else {
                        $cells[$index] = $cell->getFormattedValue();
                    }
                } catch (Throwable $exception) {
                    $cells[$index] = $cell->getValue();
                }
            }

            $fields = [];
            foreach ($headers->getHeaders() as $key => $header) {
                $fields[$key] = $header->formatValue($cells[$header->getIndex()] ?? null);
            }

            yield $row->getRowIndex() => $fields;
        }
    }

    public function eachWithCheck(callable $callback, $maxRowCount = null): SheetParser
    {
        foreach ($this->getDataIterator($maxRowCount) as $index => $fields) {
            $results = $callback($fields, $index);

            /** 中断 */
            if (false === $results) {
                break;
            }

            /** 错误 */
            if (is_array($results)) {
                $this->errors[$index] = [];
                foreach ($results as $key => $message) {
                    $this->errors[$index][$key] = $message;
                }
            }
        }

        return $this;
    }

    /**
     * @throws Exception
     */
    public function dumpErrors()
    {
        $headers = $this->getHeaders();
        if (null == $headers) {
            return;
        }

        $maxHeaderIndex = $this->getHeaders()->getMaxHeaderIndex();

        foreach ($this->getErrors() as $rowIndex => $errors) {
            /** 给错误行标色 */
            $this->getSheet()
                ->getStyle(new CellRange(
                    CellAddress::fromCellAddress(sprintf('%s%s', 'A', $rowIndex)),
                    CellAddress::fromCellAddress(sprintf('%s%s', $maxHeaderIndex, $rowIndex))
                ))
                ->getFill()->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setARGB(Color::COLOR_YELLOW);

            /** 错误单元格标色 */
            foreach ($errors as $key => $message) {
                $header = $headers->getHeaders()[$key] ?? null;
                if (null === $header) {
                    continue;
                }

                /** 设置备注 */
                if (!empty($message)) {
                    $this->getSheet()
                        ->getComment(
                            CellAddress::fromCellAddress(sprintf('%s%s', $header->getIndex(), $rowIndex))
                        )
                        ->getText()->createTextRun($message)
                        ->getFont()->setBold(true);
                }

                /** 设置颜色 */
                $this->getSheet()
                    ->getStyle(
                        CellAddress::fromCellAddress(sprintf('%s%s', $header->getIndex(), $rowIndex))
                    )
                    ->getFill()->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setARGB(Color::COLOR_RED);
            }
        }
    }
}
