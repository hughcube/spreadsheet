<?php

namespace HughCube\Spreadsheet\Tests;

use DateTime;
use HughCube\Spreadsheet\SheetParser;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class SheetParserTest extends TestCase
{
    protected function getTestFile(): string
    {
        return __DIR__ . '/resources/test.xlsx';
    }

    protected function createPatterns(): array
    {
        return [
            'index' => [
                'is' => function ($value) {
                    return '序号' === $value;
                },
                'format' => function ($value) {
                    return (int) $value;
                },
            ],
            'company' => [
                'is' => function ($value) {
                    return '受聘企业' === $value;
                },
            ],
            'name' => [
                'is' => function ($value) {
                    return '姓名' === $value;
                },
            ],
            'id_code' => [
                'is' => function ($value) {
                    return '身份证号' === $value;
                },
                'format' => function ($value) {
                    return trim(strtoupper((string) $value), '\'');
                },
            ],
            'phone' => [
                'is' => function ($value) {
                    return '手机号' === $value;
                },
            ],
            'work' => [
                'is' => function ($value) {
                    return is_string($value) && preg_match("/岗位名称/", $value);
                },
            ],
            'code' => [
                'is' => function ($value) {
                    return '原证书编号' === $value;
                },
            ],
        ];
    }

    /**
     * 测试基本的 header 解析
     */
    public function testGetHeaders()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $headers = $parser->getHeaders();

        $this->assertNotNull($headers);
        $this->assertSame(2, $headers->getIndex());
        $this->assertArrayHasKey('index', $headers->getHeaders());
        $this->assertArrayHasKey('name', $headers->getHeaders());
        $this->assertArrayHasKey('id_code', $headers->getHeaders());
        $this->assertArrayHasKey('work', $headers->getHeaders());
        $this->assertArrayHasKey('code', $headers->getHeaders());

        // 验证列映射正确
        $this->assertSame('A', $headers->getHeaders()['index']->getIndex());
        $this->assertSame('C', $headers->getHeaders()['name']->getIndex());
        $this->assertSame('D', $headers->getHeaders()['id_code']->getIndex());
        $this->assertSame('G', $headers->getHeaders()['code']->getIndex());
    }

    /**
     * 测试获取最大列
     */
    public function testGetMaxColumn()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $headers = $parser->getHeaders();

        $this->assertNotNull($headers);
        $this->assertSame('G', $headers->getMaxColumn());
    }

    /**
     * 测试数据迭代器
     */
    public function testGetDataIterator()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $rows = iterator_to_array($parser->getDataIterator());

        // test.xlsx 有 5 条数据 (第3-7行)
        $this->assertCount(5, $rows);

        // 验证第一条数据
        $first = reset($rows);
        $this->assertSame(1, $first['index']); // format 将值转为 int
        $this->assertSame('张三', $first['name']);
        $this->assertSame('320219', $first['id_code']);
        $this->assertSame('1700935T', $first['code']);
    }

    /**
     * 测试 maxRowCount 限制
     */
    public function testGetDataIteratorWithMaxRowCount()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $rows = iterator_to_array($parser->getDataIterator(2));

        $this->assertCount(2, $rows);
    }

    /**
     * 测试 eachWithCheck 空数组不记录错误
     */
    public function testEachWithCheckEmptyArrayNotError()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $parser->eachWithCheck(function ($fields, $index) {
            return [];
        });

        $this->assertEmpty($parser->getErrors());
    }

    /**
     * 测试 eachWithCheck null 不记录错误
     */
    public function testEachWithCheckNullNotError()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $parser->eachWithCheck(function ($fields, $index) {
            return null;
        });

        $this->assertEmpty($parser->getErrors());
    }

    /**
     * 测试 eachWithCheck 正确记录错误
     */
    public function testEachWithCheckRecordsErrors()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $parser->eachWithCheck(function ($fields, $index) {
            if ($fields['index'] === 1) {
                return ['name' => '姓名不能为空', 'id_code' => '身份证号格式错误'];
            }
            return null;
        });

        $errors = $parser->getErrors();
        $this->assertCount(1, $errors);

        // 第 3 行（第一条数据行）
        $firstError = reset($errors);
        $this->assertSame('姓名不能为空', $firstError['name']);
        $this->assertSame('身份证号格式错误', $firstError['id_code']);
    }

    /**
     * 测试 eachWithCheck 用 false 中断
     */
    public function testEachWithCheckBreak()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $processedCount = 0;
        $parser->eachWithCheck(function ($fields, $index) use (&$processedCount) {
            $processedCount++;
            if ($processedCount >= 2) {
                return false;
            }
            return null;
        });

        $this->assertSame(2, $processedCount);
    }

    /**
     * 测试 dumpErrors 标色和备注
     */
    public function testDumpErrors()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, $this->createPatterns());
        $parser->eachWithCheck(function ($fields, $index) {
            if ($fields['index'] === 1) {
                return ['name' => '错误信息'];
            }
            return null;
        });

        $parser->dumpErrors();

        // 检查第 3 行 A3 被标黄
        $fillA3 = $sheet->getStyle('A3')->getFill();
        $this->assertSame(Fill::FILL_SOLID, $fillA3->getFillType());
        $this->assertSame(Color::COLOR_YELLOW, $fillA3->getStartColor()->getARGB());

        // 检查第 3 行 C3 (name 列) 被标红
        $fillC3 = $sheet->getStyle('C3')->getFill();
        $this->assertSame(Fill::FILL_SOLID, $fillC3->getFillType());
        $this->assertSame(Color::COLOR_RED, $fillC3->getStartColor()->getARGB());

        // 检查备注
        $comment = $sheet->getComment('C3');
        $this->assertStringContainsString('错误信息', $comment->getText()->getPlainText());
    }

    /**
     * 测试 headers 找不到时返回 null
     */
    public function testGetHeadersReturnsNullWhenNotFound()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, [
            'nonexistent' => [
                'is' => function ($value) {
                    return '不存在的列' === $value;
                },
                'required' => true,
            ],
        ]);

        $this->assertNull($parser->getHeaders());
    }

    /**
     * 测试 getClosestHeaders 在全部匹配失败时仍返回最近匹配
     */
    public function testGetClosestHeaders()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, [
            'index' => [
                'is' => function ($value) {
                    return '序号' === $value;
                },
                'required' => true,
            ],
            'nonexistent' => [
                'is' => function ($value) {
                    return '不存在的列' === $value;
                },
                'required' => true,
            ],
        ]);

        $this->assertNull($parser->getHeaders());
        $closest = $parser->getClosestHeaders();
        $this->assertNotNull($closest);
        $this->assertArrayHasKey('index', $closest->getHeaders());
    }

    /**
     * 测试可选 header
     */
    public function testOptionalHeaders()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, [
            'index' => [
                'is' => function ($value) {
                    return '序号' === $value;
                },
                'required' => true,
            ],
            'optional_field' => [
                'is' => function ($value) {
                    return '不存在的可选列' === $value;
                },
                'required' => false,
            ],
        ]);

        $headers = $parser->getHeaders();
        $this->assertNotNull($headers);
        $this->assertArrayHasKey('index', $headers->getHeaders());
        $this->assertArrayNotHasKey('optional_field', $headers->getHeaders());
    }

    /**
     * 测试 headers 为 null 时 getDataIterator 不报错
     */
    public function testDataIteratorWithNoHeaders()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, [
            'nonexistent' => [
                'is' => function ($value) {
                    return false;
                },
                'required' => true,
            ],
        ]);

        $rows = iterator_to_array($parser->getDataIterator());
        $this->assertEmpty($rows);
    }

    /**
     * 测试 headers 为 null 时 dumpErrors 不报错
     */
    public function testDumpErrorsWithNoHeaders()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet, [
            'nonexistent' => [
                'is' => function ($value) {
                    return false;
                },
                'required' => true,
            ],
        ]);

        // 不应抛出异常
        $parser->dumpErrors();
        $this->assertTrue(true);
    }

    /**
     * 测试 addHeaderPattern 会清除缓存
     */
    public function testAddHeaderPatternInvalidatesCache()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        // 第一次: 只有一个不存在的 pattern
        $parser = SheetParser::parse($sheet, [
            'nonexistent' => [
                'is' => function ($value) {
                    return '不存在的列' === $value;
                },
                'required' => true,
            ],
        ]);

        $this->assertNull($parser->getHeaders());

        // 添加一个存在的 pattern 并移除旧的
        $parser = SheetParser::parse($sheet);
        $parser->addHeaderPattern('index', function ($value) {
            return '序号' === $value;
        });

        $headers = $parser->getHeaders();
        $this->assertNotNull($headers);
        $this->assertArrayHasKey('index', $headers->getHeaders());

        // 继续添加后重新解析
        $parser->addHeaderPattern('name', function ($value) {
            return '姓名' === $value;
        });

        $headers = $parser->getHeaders();
        $this->assertNotNull($headers);
        $this->assertArrayHasKey('name', $headers->getHeaders());
    }

    /**
     * 测试 getSheet 返回正确的 worksheet
     */
    public function testGetSheet()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = SheetParser::parse($sheet);
        $this->assertSame($sheet, $parser->getSheet());
    }

    /**
     * 测试 getHeaderPatterns
     */
    public function testGetHeaderPatterns()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $patterns = $this->createPatterns();
        $parser = SheetParser::parse($sheet, $patterns);
        $this->assertSame($patterns, $parser->getHeaderPatterns());
    }

    /**
     * 使用动态创建的表格测试多字母列 (AA, AB 等)
     */
    public function testMultiLetterColumns()
    {
        $excel = new Spreadsheet();
        $sheet = $excel->getActiveSheet();

        // 在第 1 行写 30 个 header（A~AD）
        for ($i = 0; $i < 30; $i++) {
            $col = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i + 1);
            $sheet->setCellValue($col . '1', 'col_' . $i);
        }

        // 在第 2 行写数据
        for ($i = 0; $i < 30; $i++) {
            $col = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i + 1);
            $sheet->setCellValue($col . '2', 'val_' . $i);
        }

        $parser = SheetParser::parse($sheet, [
            'first' => [
                'is' => function ($v) { return $v === 'col_0'; },
            ],
            'last' => [
                'is' => function ($v) { return $v === 'col_29'; },
            ],
        ]);

        $headers = $parser->getHeaders();
        $this->assertNotNull($headers);

        // col_29 在 AD 列
        $this->assertSame('AD', $headers->getHeaders()['last']->getIndex());
        // getMaxColumn 应该返回 AD, 不是 A
        $this->assertSame('AD', $headers->getMaxColumn());

        // 数据读取正确
        $rows = iterator_to_array($parser->getDataIterator());
        $this->assertCount(1, $rows);
        $first = reset($rows);
        $this->assertSame('val_0', $first['first']);
        $this->assertSame('val_29', $first['last']);
    }

    /**
     * 测试 header 在非第一行的情况
     */
    public function testHeaderNotOnFirstRow()
    {
        $excel = new Spreadsheet();
        $sheet = $excel->getActiveSheet();

        // 第 1-2 行是标题/空行
        $sheet->setCellValue('A1', '这是一个表格标题');
        $sheet->setCellValue('A2', '');

        // 第 3 行是表头
        $sheet->setCellValue('A3', '姓名');
        $sheet->setCellValue('B3', '年龄');

        // 第 4 行是数据
        $sheet->setCellValue('A4', '张三');
        $sheet->setCellValue('B4', '25');

        $parser = SheetParser::parse($sheet, [
            'name' => [
                'is' => function ($v) { return $v === '姓名'; },
            ],
            'age' => [
                'is' => function ($v) { return $v === '年龄'; },
                'format' => function ($v) { return (int) $v; },
            ],
        ]);

        $headers = $parser->getHeaders();
        $this->assertNotNull($headers);
        $this->assertSame(3, $headers->getIndex());

        $rows = iterator_to_array($parser->getDataIterator());
        $this->assertCount(1, $rows);
        $data = reset($rows);
        $this->assertSame('张三', $data['name']);
        $this->assertSame(25, $data['age']);
    }

    /**
     * 测试表头行后面没有数据
     */
    public function testNoDataRows()
    {
        $excel = new Spreadsheet();
        $sheet = $excel->getActiveSheet();

        $sheet->setCellValue('A1', '姓名');
        $sheet->setCellValue('B1', '年龄');

        $parser = SheetParser::parse($sheet, [
            'name' => [
                'is' => function ($v) { return $v === '姓名'; },
            ],
        ]);

        $headers = $parser->getHeaders();
        $this->assertNotNull($headers);

        $rows = iterator_to_array($parser->getDataIterator());
        $this->assertEmpty($rows);
    }

    /**
     * 测试 eachWithCheck 混合场景: 有错误, 无错误, 中断
     */
    public function testEachWithCheckMixed()
    {
        $excel = new Spreadsheet();
        $sheet = $excel->getActiveSheet();

        $sheet->setCellValue('A1', '姓名');
        $sheet->setCellValue('B1', '年龄');
        $sheet->setCellValue('A2', '张三');
        $sheet->setCellValue('B2', '25');
        $sheet->setCellValue('A3', '');
        $sheet->setCellValue('B3', '-1');
        $sheet->setCellValue('A4', '王五');
        $sheet->setCellValue('B4', '30');

        $parser = SheetParser::parse($sheet, [
            'name' => [
                'is' => function ($v) { return $v === '姓名'; },
            ],
            'age' => [
                'is' => function ($v) { return $v === '年龄'; },
            ],
        ]);

        $parser->eachWithCheck(function ($fields, $index) {
            $errors = [];
            if (empty($fields['name'])) {
                $errors['name'] = '姓名不能为空';
            }
            if ((int) $fields['age'] < 0) {
                $errors['age'] = '年龄不能为负数';
            }
            return empty($errors) ? null : $errors;
        });

        $errors = $parser->getErrors();
        // 只有第 3 行有错误
        $this->assertCount(1, $errors);
        $this->assertArrayHasKey(3, $errors);
        $this->assertSame('姓名不能为空', $errors[3]['name']);
        $this->assertSame('年龄不能为负数', $errors[3]['age']);
    }

    /**
     * 测试 dumpErrors 在动态表格上
     */
    public function testDumpErrorsOnDynamicSheet()
    {
        $excel = new Spreadsheet();
        $sheet = $excel->getActiveSheet();

        $sheet->setCellValue('A1', '姓名');
        $sheet->setCellValue('B1', '年龄');
        $sheet->setCellValue('A2', '张三');
        $sheet->setCellValue('B2', '25');
        $sheet->setCellValue('A3', '');
        $sheet->setCellValue('B3', '-1');

        $parser = SheetParser::parse($sheet, [
            'name' => [
                'is' => function ($v) { return $v === '姓名'; },
            ],
            'age' => [
                'is' => function ($v) { return $v === '年龄'; },
            ],
        ]);

        $parser->eachWithCheck(function ($fields, $index) {
            if (empty($fields['name'])) {
                return ['name' => '必填'];
            }
            return null;
        });

        $parser->dumpErrors();

        // 第 3 行被标黄
        $fillA3 = $sheet->getStyle('A3')->getFill();
        $this->assertSame(Fill::FILL_SOLID, $fillA3->getFillType());

        // name 列 (A3) 被标红
        $this->assertSame(
            Color::COLOR_RED,
            $sheet->getStyle('A3')->getFill()->getStartColor()->getARGB()
        );

        // 第 2 行不受影响
        $this->assertNotSame(
            Color::COLOR_YELLOW,
            $sheet->getStyle('A2')->getFill()->getStartColor()->getARGB()
        );
    }

    /**
     * 测试 ParseSheet 向后兼容
     */
    public function testParseSheetBackwardCompatibility()
    {
        $excel = IOFactory::load($this->getTestFile());
        $sheet = $excel->getActiveSheet();

        $parser = \HughCube\Spreadsheet\ParseSheet::parse($sheet, [
            'index' => [
                'is' => function ($value) {
                    return '序号' === $value;
                },
            ],
        ]);

        $this->assertInstanceOf(\HughCube\Spreadsheet\SheetParser::class, $parser);
        $this->assertNotNull($parser->getHeaders());
    }

    /**
     * 测试同一行中不会将一个 pattern 匹配到多列
     */
    public function testPatternMatchesOnlyOnce()
    {
        $excel = new Spreadsheet();
        $sheet = $excel->getActiveSheet();

        // 两列都包含 "序号"
        $sheet->setCellValue('A1', '序号');
        $sheet->setCellValue('B1', '序号');
        $sheet->setCellValue('C1', '姓名');
        $sheet->setCellValue('A2', '1');
        $sheet->setCellValue('B2', '2');
        $sheet->setCellValue('C2', '张三');

        $parser = SheetParser::parse($sheet, [
            'index' => [
                'is' => function ($v) { return $v === '序号'; },
            ],
            'name' => [
                'is' => function ($v) { return $v === '姓名'; },
            ],
        ]);

        $headers = $parser->getHeaders();
        $this->assertNotNull($headers);
        // "序号" 只能匹配到第一个出现的列 A
        $this->assertSame('A', $headers->getHeaders()['index']->getIndex());

        $rows = iterator_to_array($parser->getDataIterator());
        $first = reset($rows);
        /** PhpSpreadsheet 把字符串 '1' 推断为数字 cell, 修复后数字 cell 返回原始 int (不再是字符串 '1') */
        $this->assertSame(1, $first['index']);
        $this->assertSame('张三', $first['name']);
    }

    /**
     * 回归测试: 数字单元格的 numberFormat 不能把原始数字截断
     *
     * 历史 bug: getFormattedValue / rangeToArray(formatData=true) 会按 cell 的
     * numberFormat 把数字渲染成 "1,000" / "￥1,000" / "50%" 等字符串, 调用方
     * floatval 时会从第一个非数字字符截断 → 1000 变 1, 百分比 50% 变 0 等等.
     *
     * 修复后: 数字 / 公式且非日期单元格应直接返回原始数字; 日期仍返回字符串;
     * 字符串 / 布尔 / 空值与之前一致.
     *
     * @see docs/fix-numeric-format-truncation.md
     */
    public function testNumericFormatNotTruncated(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->setCellValue('A1', '金额');
        $sheet->setCellValue('B1', '货币');
        $sheet->setCellValue('C1', '百分比');
        $sheet->setCellValue('D1', '日期');
        $sheet->setCellValue('E1', '公式');
        $sheet->setCellValue('F1', '普通数字');
        $sheet->setCellValue('G1', '字符串');

        /** 千分位: numFormatId=3 内置格式, 之前会被渲染成 "1,000" */
        $sheet->setCellValue('A2', 1000);
        $sheet->getStyle('A2')->getNumberFormat()->setFormatCode('#,##0');

        /** 货币: 渲染后是 "¥1,234.56", floatval 拿到 0 */
        $sheet->setCellValue('B2', 1234.56);
        $sheet->getStyle('B2')->getNumberFormat()->setFormatCode('"¥"#,##0.00');

        /** 百分比: 内部 0.5, 渲染后 "50%", floatval 拿到 0.0 */
        $sheet->setCellValue('C2', 0.5);
        $sheet->getStyle('C2')->getNumberFormat()->setFormatCode('0%');

        /** 日期: 内部是 Excel serial number, 必须保留 getFormattedValue 路径 */
        $sheet->setCellValue('D2', Date::PHPToExcel(new DateTime('2026-05-27 15:30:00')));
        $sheet->getStyle('D2')->getNumberFormat()->setFormatCode('yyyy-mm-dd hh:mm:ss');

        /** 公式: 计算结果是数字, 之前 getFormattedValue 会带千分位 */
        $sheet->setCellValue('E2', '=A2*2');
        $sheet->getStyle('E2')->getNumberFormat()->setFormatCode('#,##0');

        /** 普通数字: 无 numberFormat, 不应受影响 */
        $sheet->setCellValue('F2', 42);

        /** 字符串: 不应受影响 */
        $sheet->setCellValue('G2', 'hello');

        $parser = SheetParser::parse($sheet)
            ->addHeaderPattern('amount', fn($v) => '金额' === $v)
            ->addHeaderPattern('currency', fn($v) => '货币' === $v)
            ->addHeaderPattern('percent', fn($v) => '百分比' === $v)
            ->addHeaderPattern('date', fn($v) => '日期' === $v)
            ->addHeaderPattern('formula', fn($v) => '公式' === $v)
            ->addHeaderPattern('plain', fn($v) => '普通数字' === $v)
            ->addHeaderPattern('text', fn($v) => '字符串' === $v);

        $rows = iterator_to_array($parser->getDataIterator());
        $first = reset($rows);

        /** 核心断言: 数字 cell 不应被千分位 / 货币 / 百分比格式截断 */
        $this->assertSame(1000, $first['amount'], '千分位数字应返回原始 1000, 不是 "1,000"');
        $this->assertSame(1000.0, floatval($first['amount']), 'floatval 不应被截断成 1.0');

        $this->assertEqualsWithDelta(1234.56, $first['currency'], 0.001, '货币数字应返回原始 float');
        $this->assertEqualsWithDelta(1234.56, floatval($first['currency']), 0.001);

        $this->assertSame(0.5, $first['percent'], '百分比应返回 0~1 表示的数字, 不是 "50%"');

        /** 日期必须仍是字符串形式 */
        $this->assertIsString($first['date'], '日期单元格应仍返回格式化字符串');
        $this->assertStringContainsString('2026-05-27', $first['date']);

        /** 公式应返回原始计算结果, 不带千分位 */
        $this->assertEqualsWithDelta(2000, $first['formula'], 0.001);
        $this->assertNotSame('2,000', $first['formula']);

        /** 普通无格式数字 */
        $this->assertSame(42, $first['plain']);

        /** 字符串单元格行为不变 */
        $this->assertSame('hello', $first['text']);
    }
}
