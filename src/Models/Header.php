<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/5
 * Time: 23:26
 */

namespace HughCube\Spreadsheet\Models;

/**
 * 表示电子表格中的单个列头定义
 *
 * 包含列的位置索引(列字母)、标题文本以及可选的值格式化函数
 */
class Header
{
    /**
     * 列索引, 即列字母标识, 如 A, B, AA 等
     *
     * @var string
     */
    protected string $index;

    /**
     * 列标题文本, 即表头单元格中的内容
     *
     * @var string|null
     */
    protected string|null $title;

    /**
     * 值格式化回调函数, 用于转换该列的单元格数据
     *
     * @var callable|null
     */
    protected mixed $format;

    /**
     * @param string        $index  列字母索引, 如 'A', 'B', 'AA'
     * @param string|null   $title  表头单元格的文本内容
     * @param callable|null $format 可选的值格式化函数, 签名: function($value): mixed
     */
    public function __construct(string $index, string|null $title, ?callable $format = null)
    {
        $this->index = $index;
        $this->title = $title;
        $this->format = $format;
    }

    /**
     * 获取列字母索引
     *
     * @return string 列字母, 如 'A', 'B', 'AA'
     */
    public function getIndex(): string
    {
        return $this->index;
    }

    /**
     * 获取表头标题文本
     *
     * @return string|null
     */
    public function getTitle(): string|null
    {
        return $this->title;
    }

    /**
     * 使用格式化函数处理单元格值
     *
     * 如果未设置格式化函数, 则原样返回值
     *
     * @param mixed $value 原始单元格值
     * @return mixed 格式化后的值
     */
    public function formatValue(mixed $value): mixed
    {
        return null === $this->format ? $value : call_user_func($this->format, $value);
    }
}
