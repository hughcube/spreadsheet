<?php

namespace HughCube\Spreadsheet\Exporter;

interface DriverInterface
{
    /**
     * 打开文件准备写入
     */
    public function open(string $filename): void;

    /**
     * 添加工作表
     */
    public function addSheet(string $name): void;

    /**
     * 写入一行数据, 行列索引均为 0-based
     *
     * @param int $rowIndex 行索引 (0-based)
     * @param array $values 单元格值, key 为 0-based 列索引
     * @param array<int, Style> $cellStyles 各单元格已解析的最终样式, key 为 0-based 列索引
     */
    public function writeRow(int $rowIndex, array $values, array $cellStyles = []): void;

    /**
     * 设置列宽
     *
     * @param int $colIndex 0-based 列索引
     */
    public function setColumnWidth(int $colIndex, float $width): void;

    /**
     * 设置行高
     *
     * @param int $rowIndex 0-based 行索引
     */
    public function setRowHeight(int $rowIndex, float $height): void;

    /**
     * 合并单元格, 所有参数均为 0-based
     */
    public function mergeCells(int $startRow, int $startCol, int $endRow, int $endCol): void;

    /**
     * 关闭并保存文件
     */
    public function close(): void;

    /**
     * 当前驱动是否支持样式
     */
    public function supportsStyle(): bool;

    /**
     * 获取底层原生实例
     *
     * PhpSpreadsheet: 返回 \PhpOffice\PhpSpreadsheet\Spreadsheet
     * OpenSpout:      返回 \OpenSpout\Writer\XLSX\Writer
     * Xlswriter:      返回 \Vtiful\Kernel\Excel
     * Csv:            返回 resource (文件句柄)
     *
     * @return mixed
     */
    public function getNativeInstance(): mixed;
}
