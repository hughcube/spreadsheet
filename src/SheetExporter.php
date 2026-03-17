<?php

namespace HughCube\Spreadsheet;

use HughCube\Spreadsheet\Exporter\Coordinate;
use HughCube\Spreadsheet\Exporter\DriverInterface;
use HughCube\Spreadsheet\Exporter\Drivers\CsvDriver;
use HughCube\Spreadsheet\Exporter\Drivers\OpenSpoutDriver;
use HughCube\Spreadsheet\Exporter\Drivers\PhpSpreadsheetDriver;
use HughCube\Spreadsheet\Exporter\Drivers\XlswriterDriver;
use HughCube\Spreadsheet\Exporter\Style;
use InvalidArgumentException;
use LogicException;

class SheetExporter
{
    protected DriverInterface $driver;

    /** 文件是否已打开 */
    protected bool $isOpen = false;

    /** 延迟打开的文件路径 */
    protected ?string $pendingFilename = null;

    /** 当前 sheet 索引, -1 表示尚未添加 sheet */
    protected int $currentSheet = -1;

    /** @var array<int, Style> 表级默认样式 (常驻内存, 极小) */
    protected array $defaultStyles = [];

    /** @var array<int, array<int, Style>> 列样式 [sheetIdx => [colIdx => Style]] (常驻内存, 极小) */
    protected array $columnStyles = [];

    /** @var array<int, array<int, Style>> 预设行样式 [sheetIdx => [rowIdx => Style]], 写入后清除 */
    protected array $rowStyles = [];

    /** @var array<int, array<int, array<int, Style>>> 预设单元格样式, 写入后清除 */
    protected array $cellStyles = [];

    /** @var array<int, array<int, float>> 预设行高, 写入后清除 */
    protected array $rowHeights = [];

    /** @var array<int, array<array{0:int,1:int,2:int,3:int}>> 合并单元格坐标, 切换sheet/close时应用 */
    protected array $mergeCells = [];

    /** @var array<int, int> 每个 sheet 的当前行计数器 */
    protected array $rowCounters = [];

    /**
     * 创建导出器, 自动选择最优驱动或指定驱动
     *
     * @param string|DriverInterface|null $driver 驱动名称、驱动实例或 null(自动选择)
     *   支持: 'auto', 'csv', 'phpspreadsheet', 'openspout', 'xlswriter'
     * @param string|null $filename 文件路径, 传入后 addSheet() 时自动 open
     */
    public static function create(string|DriverInterface|null $driver = null, ?string $filename = null): static
    {
        if ($driver instanceof DriverInterface) {
            /** @phpstan-ignore new.static */
            $instance = new static();
            $instance->driver = $driver;
            $instance->pendingFilename = $filename;
            return $instance;
        }

        if (null === $driver || 'auto' === $driver) {
            $driver = static::detectBestDriver();
        }

        /** @phpstan-ignore new.static */
        $instance = new static();
        $instance->driver = match ($driver) {
            'csv' => new CsvDriver(),
            'phpspreadsheet' => new PhpSpreadsheetDriver(),
            'openspout', 'spout' => new OpenSpoutDriver(),
            'xlswriter' => new XlswriterDriver(),
            default => throw new InvalidArgumentException("未知驱动: {$driver}"),
        };
        $instance->pendingFilename = $filename;

        return $instance;
    }

    /**
     * 自动检测可用的最优驱动
     *
     * 优先级: xlswriter(C扩展最快) > openspout(流式低内存) > phpspreadsheet(功能最全) > csv
     */
    protected static function detectBestDriver(): string
    {
        if (extension_loaded('xlswriter')) {
            return 'xlswriter';
        }

        if (class_exists(\OpenSpout\Writer\XLSX\Writer::class)) {
            return 'openspout';
        }

        if (class_exists(\PhpOffice\PhpSpreadsheet\Spreadsheet::class)) {
            return 'phpspreadsheet';
        }

        return 'csv';
    }

    public function getDriver(): DriverInterface
    {
        return $this->driver;
    }

    /**
     * 打开文件, 准备写入
     */
    public function openFile(string $filename): static
    {
        if ($this->isOpen) {
            throw new LogicException('文件已打开, 请先调用 close()');
        }

        $this->driver->open($filename);
        $this->isOpen = true;
        return $this;
    }

    /**
     * 添加工作表
     *
     * 如果文件未打开但 create() 时传了 filename, 会自动 open.
     * 如果当前有活跃 sheet, 会先完成该 sheet 的收尾工作(应用合并单元格等).
     *
     * @param string|null $name sheet 名称, null 则自动生成 Sheet1, Sheet2, ...
     */
    public function addSheet(?string $name = null): static
    {
        /** 自动 open */
        if (!$this->isOpen && null !== $this->pendingFilename) {
            $this->openFile($this->pendingFilename);
            $this->pendingFilename = null;
        }

        $this->ensureOpen();

        /** 收尾上一个 sheet */
        if ($this->currentSheet >= 0) {
            $this->finalizeCurrentSheet();
        }

        $this->currentSheet++;
        $this->rowCounters[$this->currentSheet] = 0;

        $sheetName = $name ?? ('Sheet' . ($this->currentSheet + 1));
        $this->driver->addSheet($sheetName);
        return $this;
    }

    /**
     * 设置当前 sheet 的默认样式(表级)
     *
     * 应在 addSheet() 之后, writeRow() 之前调用
     */
    public function setDefaultStyle(Style $style): static
    {
        $this->defaultStyles[$this->currentSheet] = $style;
        return $this;
    }

    /**
     * 设置列样式
     *
     * @param string $column 列字母, 如 'A', 'B', 'AA'
     */
    public function setColumnStyle(string $column, Style $style): static
    {
        $colIdx = Coordinate::columnLetterToIndex($column);
        $this->columnStyles[$this->currentSheet][$colIdx] = $style;
        return $this;
    }

    /**
     * 预设行样式(在写入该行之前调用)
     *
     * @param int $row 1-based 行号
     */
    public function setRowStyle(int $row, Style $style): static
    {
        $this->rowStyles[$this->currentSheet][$row - 1] = $style;
        return $this;
    }

    /**
     * 预设单元格样式(在写入该行之前调用)
     *
     * @param string $cell 单元格地址, 如 'A1', 'B3'
     */
    public function setCellStyle(string $cell, Style $style): static
    {
        [$row, $col] = Coordinate::parseCellAddress($cell);
        $this->cellStyles[$this->currentSheet][$row][$col] = $style;
        return $this;
    }

    /**
     * 设置列宽, 立即写入驱动
     *
     * @param string $column 列字母
     */
    public function setColumnWidth(string $column, float $width): static
    {
        $this->ensureOpen();
        $this->driver->setColumnWidth(Coordinate::columnLetterToIndex($column), $width);
        return $this;
    }

    /**
     * 预设行高(在写入该行之前调用)
     *
     * @param int $row 1-based 行号
     */
    public function setRowHeight(int $row, float $height): static
    {
        $this->rowHeights[$this->currentSheet][$row - 1] = $height;
        return $this;
    }

    /**
     * 写入一行数据, 立即落盘, 自动递增行号
     *
     * @param array $values 单元格值
     * @param Style|null $rowStyle 该行的行级样式(内联)
     */
    public function writeRow(array $values, ?Style $rowStyle = null): static
    {
        $this->ensureOpen();

        $sheetIdx = $this->currentSheet;
        $rowIdx = $this->rowCounters[$sheetIdx]++;
        $values = array_values($values);

        /** 行高 */
        if (isset($this->rowHeights[$sheetIdx][$rowIdx])) {
            $this->driver->setRowHeight($rowIdx, $this->rowHeights[$sheetIdx][$rowIdx]);
            unset($this->rowHeights[$sheetIdx][$rowIdx]);
        }

        /** 解析每个单元格的最终样式 */
        $cellStyles = [];
        foreach ($values as $colIdx => $value) {
            $resolved = $this->resolveStyle($sheetIdx, $rowIdx, $colIdx, $rowStyle);
            if (null !== $resolved) {
                $cellStyles[$colIdx] = $resolved;
            }
        }

        /** 立即写入驱动 */
        $this->driver->writeRow($rowIdx, $values, $cellStyles);

        /** 释放已消费的预设样式, 避免内存堆积 */
        unset($this->rowStyles[$sheetIdx][$rowIdx]);
        unset($this->cellStyles[$sheetIdx][$rowIdx]);

        return $this;
    }

    /**
     * 批量写入多行数据
     *
     * @param iterable $rows 行数据集合, 支持数组和 Generator
     * @param Style|null $rowStyle 所有行的行级样式
     */
    public function writeRows(iterable $rows, ?Style $rowStyle = null): static
    {
        foreach ($rows as $row) {
            $this->writeRow($row, $rowStyle);
        }
        return $this;
    }

    /**
     * 追加一行, writeRow 的别名
     */
    public function appendRow(array $values, ?Style $rowStyle = null): static
    {
        return $this->writeRow($values, $rowStyle);
    }

    /**
     * 批量追加多行数据
     */
    public function appendRows(iterable $rows, ?Style $rowStyle = null): static
    {
        return $this->writeRows($rows, $rowStyle);
    }

    /**
     * 合并单元格(坐标缓存, 切换 sheet 或 close 时应用)
     *
     * @param string $range 范围, 如 'A1:C1', 'A1:A3'
     */
    public function mergeCells(string $range): static
    {
        $this->mergeCells[$this->currentSheet][] = Coordinate::parseRange($range);
        return $this;
    }

    /**
     * 直接操作原生实例
     *
     * @param callable $callback function(mixed $nativeInstance): void
     */
    public function tap(callable $callback): static
    {
        $callback($this->driver->getNativeInstance());
        return $this;
    }

    /**
     * 关闭并保存文件
     *
     * @param callable|null $beforeClose 关闭前的回调, 接收原生实例
     */
    public function close(?callable $beforeClose = null): void
    {
        if (!$this->isOpen) {
            return;
        }

        /** 收尾最后一个 sheet */
        if ($this->currentSheet >= 0) {
            $this->finalizeCurrentSheet();
        }

        if (null !== $beforeClose) {
            $beforeClose($this->driver->getNativeInstance());
        }

        $this->driver->close();
        $this->isOpen = false;
    }

    /**
     * 解析某个单元格的最终样式
     *
     * 优先级: 单元格 > 内联行样式 > 预设行样式 > 列 > 表默认
     */
    protected function resolveStyle(int $sheetIdx, int $rowIdx, int $colIdx, ?Style $inlineRowStyle): ?Style
    {
        $style = $this->defaultStyles[$sheetIdx] ?? null;

        $colStyle = $this->columnStyles[$sheetIdx][$colIdx] ?? null;
        if (null !== $colStyle) {
            $style = (null !== $style) ? $style->merge($colStyle) : $colStyle;
        }

        $presetRowStyle = $this->rowStyles[$sheetIdx][$rowIdx] ?? null;
        if (null !== $presetRowStyle) {
            $style = (null !== $style) ? $style->merge($presetRowStyle) : $presetRowStyle;
        }

        if (null !== $inlineRowStyle) {
            $style = (null !== $style) ? $style->merge($inlineRowStyle) : $inlineRowStyle;
        }

        $cellStyle = $this->cellStyles[$sheetIdx][$rowIdx][$colIdx] ?? null;
        if (null !== $cellStyle) {
            $style = (null !== $style) ? $style->merge($cellStyle) : $cellStyle;
        }

        return $style;
    }

    /**
     * 完成当前 sheet 的收尾(应用合并单元格)
     */
    protected function finalizeCurrentSheet(): void
    {
        foreach ($this->mergeCells[$this->currentSheet] ?? [] as [$sr, $sc, $er, $ec]) {
            $this->driver->mergeCells($sr, $sc, $er, $ec);
        }
        unset($this->mergeCells[$this->currentSheet]);
    }

    protected function ensureOpen(): void
    {
        if (!$this->isOpen) {
            throw new LogicException('文件未打开, 请先调用 openFile()');
        }
    }
}
