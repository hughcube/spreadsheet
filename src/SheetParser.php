<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/5
 * Time: 23:30
 */

namespace HughCube\Spreadsheet;

use Generator;
use HughCube\Spreadsheet\Models\Headers;
use PhpOffice\PhpSpreadsheet\Cell\CellAddress;
use PhpOffice\PhpSpreadsheet\Cell\CellRange;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Throwable;

/**
 * 电子表格解析器
 *
 * 通过 pattern 匹配自动识别表头, 迭代数据行并进行格式化,
 * 支持数据验证、错误收集以及将错误标记回工作表(标色+备注)
 *
 * 使用示例:
 *   $parser = SheetParser::parse($worksheet, [
 *       'name' => ['is' => fn($v) => $v === '姓名', 'required' => true],
 *       'age'  => ['is' => fn($v) => $v === '年龄', 'format' => fn($v) => (int)$v],
 *   ]);
 *   $parser->eachWithCheck(function ($fields, $rowIndex) {
 *       if (empty($fields['name'])) return ['name' => '姓名不能为空'];
 *   });
 *   $parser->dumpErrors();
 */
class SheetParser
{
    /**
     * 工作表实例
     *
     * @var Worksheet
     */
    protected Worksheet $sheet;

    /**
     * 已解析的表头缓存
     *
     * false: 尚未解析, null: 解析后未找到匹配, Headers: 解析成功
     *
     * @var false|null|Headers
     */
    protected Headers|null|false $headers = false;

    /**
     * 最近匹配的表头(即使未完全满足 required 条件)
     *
     * @var null|Headers
     */
    protected null|Headers $closestHeaders = null;

    /**
     * 表头匹配规则数组
     *
     * 格式: ['key' => ['is' => callable, 'format' => callable|null, 'required' => bool]]
     *
     * @var array
     */
    protected array $headerPatterns;

    /**
     * 验证错误集合
     *
     * 格式: [行号 => ['字段key' => '错误信息', ...], ...]
     *
     * @var array<int, array<string, string>>
     */
    protected array $errors = [];

    /**
     * 创建解析器实例(工厂方法)
     *
     * @param Worksheet $sheet 要解析的工作表
     * @param array $patterns 表头匹配规则
     * @return static
     */
    public static function parse(Worksheet $sheet, array $patterns = []): static
    {
        /** @phpstan-ignore-next-line */
        return new static($sheet, $patterns);
    }

    /**
     * @param Worksheet $sheet 工作表实例
     * @param array $patterns 表头匹配规则
     */
    protected function __construct(Worksheet $sheet, array $patterns)
    {
        $this->sheet = $sheet;
        $this->headerPatterns = $patterns;
    }

    /**
     * 获取关联的工作表实例
     *
     * @return Worksheet
     */
    public function getSheet(): Worksheet
    {
        return $this->sheet;
    }

    /**
     * 获取当前配置的所有表头匹配规则
     *
     * @return array
     */
    public function getHeaderPatterns(): array
    {
        return $this->headerPatterns;
    }

    /**
     * 动态添加一个表头匹配规则
     *
     * 添加后会清除已缓存的表头解析结果, 确保下次 getHeaders() 时重新解析
     *
     * @param string $name 规则名称, 作为数据字段的 key
     * @param callable|null $is 匹配函数, 签名: function($cellValue): bool
     * @param callable|null $format 格式化函数, 签名: function($value): mixed
     * @param bool $required 是否为必需列, 默认 true
     * @return $this
     */
    public function addHeaderPattern(
        string $name,
        ?callable $is,
        ?callable $format = null,
        bool $required = true
    ): static {
        $this->headerPatterns[$name] = ['is' => $is, 'format' => $format, 'required' => $required];

        /** 清除缓存, 确保新增的 pattern 生效 */
        $this->headers = false;
        $this->closestHeaders = null;

        return $this;
    }

    /**
     * 获取解析后的表头
     *
     * 首次调用时触发表头解析, 后续调用返回缓存结果.
     * 如果所有 required 的 pattern 都匹配到, 返回 Headers 实例; 否则返回 null
     *
     * @return null|Headers
     */
    public function getHeaders(): ?Headers
    {
        if (false === $this->headers) {
            list($this->headers, $this->closestHeaders) = Headers::parse($this, $this->getHeaderPatterns());
        }
        return $this->headers;
    }

    /**
     * 获取最近匹配的表头
     *
     * 即使未找到完全匹配的表头行, 也会返回匹配度最高的结果,
     * 可用于在表头不完全匹配时给出提示信息
     *
     * @return null|Headers
     */
    public function getClosestHeaders(): ?Headers
    {
        $this->getHeaders();

        return $this->closestHeaders;
    }

    /**
     * 获取所有已收集的验证错误
     *
     * @return array<int, array<string, string>> 格式: [行号 => ['字段key' => '错误信息']]
     */
    public function getErrors(): array
    {
        return $this->errors;
    }

    /**
     * 获取数据行迭代器(生成器)
     *
     * 从表头下一行开始逐行读取数据, 根据 Header 的列映射提取字段值,
     * 并通过 Header::formatValue() 进行格式化
     *
     * @param int|null $maxRowCount 最大读取行数, null 表示读取全部数据行
     * @return Generator<int, array<string, mixed>> 键为行号, 值为字段数组 ['key' => formatted_value]
     */
    public function getDataIterator(?int $maxRowCount = null): Generator
    {
        $headers = $this->getHeaders();
        if (null !== $headers) {
            $endColumn = $headers->getMaxColumn();

            $startRowIndex = $headers->getIndex() + 1;
            $highestRow = $this->getSheet()->getHighestRow();

            $endRowIndex = $highestRow;
            if (!empty($maxRowCount)) {
                $endRowIndex = min($startRowIndex + $maxRowCount - 1, $highestRow);
            }

            if ($startRowIndex <= $highestRow) {
                for ($rowIndex = $startRowIndex; $rowIndex <= $endRowIndex; $rowIndex++) {
                    /** 通过 rangeToArray 批量读取, 避免逐个创建 Cell 对象导致内存膨胀 */
                    $range = sprintf('A%d:%s%d', $rowIndex, $endColumn, $rowIndex);
                    try {
                        $rowData = $this->getSheet()->rangeToArray($range, null, true, true, true);
                        $cells = $rowData[$rowIndex] ?? [];
                    } catch (Throwable $exception) {
                        /** 批量读取失败(含公式计算异常), 降级为逐单元格处理, 仅失败的 cell 取原始值 */
                        $cells = [];
                        foreach ($this->getSheet()->getRowIterator($rowIndex, $rowIndex) as $row) {
                            foreach ($row->getCellIterator('A', $endColumn) as $index => $cell) {
                                try {
                                    $cells[$index] = $cell->getFormattedValue();
                                } catch (Throwable $e) {
                                    $cells[$index] = $cell->getValue();
                                }
                            }
                        }
                    }

                    $fields = [];
                    foreach ($headers->getHeaders() as $key => $header) {
                        $fields[$key] = $header->formatValue($cells[$header->getIndex()] ?? null);
                    }

                    yield $rowIndex => $fields;
                }
            }
        }
    }

    /**
     * 带验证回调的数据行遍历
     *
     * 对每一行数据调用回调函数进行验证:
     * - 回调返回 null 或非数组: 该行无错误
     * - 回调返回非空数组: 该行有错误, 格式为 ['字段key' => '错误信息'], 记录到 errors
     * - 回调返回 false: 中断遍历
     * - 回调返回空数组 []: 视为无错误, 不记录
     *
     * @param callable $callback 验证回调, 签名: function(array $fields, int $rowIndex): array|false|null
     * @param int|null $maxRowCount 最大处理行数, null 表示处理全部
     * @return $this
     */
    public function eachWithCheck(callable $callback, ?int $maxRowCount = null): static
    {
        foreach ($this->getDataIterator($maxRowCount) as $index => $fields) {
            $results = $callback($fields, $index);

            /** 中断 */
            if (false === $results) {
                break;
            }

            /** 错误: 只记录非空的错误数组 */
            if (is_array($results) && !empty($results)) {
                $this->errors[$index] = [];
                foreach ($results as $key => $message) {
                    $this->errors[$index][$key] = $message;
                }
            }
        }

        return $this;
    }

    /**
     * 将收集到的错误标记到工作表中
     *
     * 对每个有错误的行:
     * - 整行背景标为黄色(COLOR_YELLOW)
     * - 有错误的具体单元格背景标为红色(COLOR_RED)
     * - 在错误单元格添加批注(Comment), 内容为错误信息
     *
     * @return void
     * @throws Exception
     */
    public function dumpErrors(): void
    {
        $headers = $this->getHeaders();
        if (null === $headers) {
            return;
        }
        $maxColumn = $headers->getMaxColumn();

        foreach ($this->getErrors() as $rowIndex => $errors) {
            /** 给错误行标色 */
            $this->getSheet()
                ->getStyle(new CellRange(
                    CellAddress::fromCellAddress(sprintf('%s%s', 'A', $rowIndex)),
                    CellAddress::fromCellAddress(sprintf('%s%s', $maxColumn, $rowIndex))
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
