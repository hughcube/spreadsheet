<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/5
 * Time: 23:26
 */

namespace HughCube\Spreadsheet\Models;

use HughCube\Spreadsheet\SheetParser;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use Throwable;

/**
 * 表头集合, 管理一组已解析的 Header 对象
 *
 * 负责从工作表中自动识别表头行, 通过 pattern 匹配确定每列的含义
 */
class Headers
{
    /**
     * 表头所在的行号
     *
     * @var int
     */
    protected int $index;

    /**
     * 已解析的表头集合, key 为 pattern 名称, value 为 Header 对象
     *
     * @var array<string, Header>
     */
    protected array $headers;

    /**
     * @param int   $index   表头所在行号
     * @param array $headers 表头集合, key => Header
     */
    public function __construct(int $index, array $headers)
    {
        $this->index = $index;
        $this->headers = $headers;
    }

    /**
     * 获取所有已解析的表头
     *
     * @return array<string, Header> key 为 pattern 名称, value 为 Header 对象
     */
    public function getHeaders(): array
    {
        return $this->headers;
    }

    /**
     * 获取表头所在的行号
     *
     * @return int
     */
    public function getIndex(): int
    {
        return $this->index;
    }

    /**
     * 获取所有表头中最大的列字母
     *
     * @deprecated 使用 getMaxColumn() 代替
     * @see self::getMaxColumn()
     * @return string
     */
    public function getMaxHeaderIndex(): string
    {
        return $this->getMaxColumn();
    }

    /**
     * 获取所有表头中最大的列字母
     *
     * 通过数值比较列索引, 正确处理多字母列名(如 AA, AB > Z)
     *
     * @return string 最大列字母, 如 'G', 'AA' 等, 无表头时返回 'A'
     */
    public function getMaxColumn(): string
    {
        $maxNumeric = Coordinate::columnIndexFromString('A');
        foreach ($this->getHeaders() as $header) {
            $numeric = Coordinate::columnIndexFromString($header->getIndex());
            if ($numeric > $maxNumeric) {
                $maxNumeric = $numeric;
            }
        }
        return Coordinate::stringFromColumnIndex($maxNumeric);
    }

    /**
     * 从工作表中解析表头行
     *
     * 逐行扫描工作表, 使用 pattern 回调匹配每个单元格, 找到包含所有必需列的行作为表头行.
     * 同时记录匹配度最高的行作为 closestHeaders, 即使未完全匹配也可使用.
     *
     * 每个 pattern 在同一行中只会匹配一次(第一个匹配的单元格), 避免重复列.
     *
     * @param SheetParser $parse       解析器实例
     * @param array       $patterns    表头匹配规则, 格式: ['key' => ['is' => callable, 'format' => callable|null, 'required' => bool]]
     * @param int         $maxScanRows 最大扫描行数, 0 表示不限制, 默认 50
     * @return array{0: static|null, 1: static|null} [完全匹配的表头|null, 最近匹配的表头|null]
     */
    public static function parse(SheetParser $parse, array $patterns, int $maxScanRows = 50): array
    {
        if (empty($patterns)) {
            return [null, null];
        }

        $requiredKeys = [];
        foreach ($patterns as $key => $pattern) {
            if ($pattern['required'] ?? true) {
                $requiredKeys[] = $key;
            }
        }

        $sheet = $parse->getSheet();
        $highestRow = $sheet->getHighestRow();
        $highestCol = $sheet->getHighestDataColumn();
        $patternCount = count($patterns);

        /** 限制扫描行数 */
        $scanEndRow = ($maxScanRows > 0) ? min($highestRow, $maxScanRows) : $highestRow;

        $match = false;
        $closestHeaders = null;

        for ($rowIndex = 1; $rowIndex <= $scanEndRow; $rowIndex++) {
            /** 通过 rangeToArray 批量读取整行, 避免逐个创建 Cell 对象导致内存膨胀 */
            $range = sprintf('A%d:%s%d', $rowIndex, $highestCol, $rowIndex);
            try {
                $rowData = $sheet->rangeToArray($range, null, true, true, true);
                $cells = $rowData[$rowIndex] ?? [];
            } catch (Throwable $exception) {
                /** 批量读取失败, 降级为逐单元格处理, 仅失败的 cell 取原始值 */
                $cells = [];
                foreach ($sheet->getRowIterator($rowIndex, $rowIndex) as $row) {
                    foreach ($row->getCellIterator('A', $highestCol) as $index => $cell) {
                        try {
                            $cells[$index] = $cell->getFormattedValue();
                        } catch (Throwable $e) {
                            $cells[$index] = $cell->getValue();
                        }
                    }
                }
            }

            /** 跳过空行 */
            $hasValue = false;
            foreach ($cells as $cell) {
                if ($cell !== null && $cell !== '') {
                    $hasValue = true;
                    break;
                }
            }
            if (!$hasValue) {
                continue;
            }

            /** 尝试解析 Header, 已匹配的 pattern 不再参与后续 cell 匹配 */
            $headers = [];
            $matchedCount = 0;
            foreach ($cells as $index => $cell) {
                /** 所有 pattern 都已匹配, 无需继续遍历 cell */
                if ($matchedCount >= $patternCount) {
                    break;
                }
                foreach ($patterns as $key => $pattern) {
                    if (isset($headers[$key])) {
                        continue;
                    }
                    if (($pattern['is'])($cell)) {
                        $headers[$key] = new Header($index, $cell, $pattern['format'] ?? null);
                        $matchedCount++;
                        break;
                    }
                }
            }

            if ($matchedCount > 0) {
                if (!$closestHeaders instanceof static || $matchedCount > count($closestHeaders->getHeaders())) {
                    /** @phpstan-ignore-next-line */
                    $closestHeaders = new static($rowIndex, $headers);
                }

                if (empty(array_diff($requiredKeys, array_keys($headers)))) {
                    /** @phpstan-ignore-next-line */
                    $closestHeaders = new static($rowIndex, $headers);
                    $match = true;
                    break;
                }
            }
        }

        return [($match ? $closestHeaders : null), $closestHeaders];
    }
}
