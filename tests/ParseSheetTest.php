<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/6
 * Time: 01:22
 */

namespace HughCube\Spreadsheet\Tests;

use DateTime;
use HughCube\Spreadsheet\ParseSheet;
use HughCube\Spreadsheet\SheetParser;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ParseSheetTest extends TestCase
{
    ///**
    // * @throws Exception
    // */
    //public function testHandle()
    //{
    //    $file = __DIR__.'/resources/test.xlsx';
    //    $excel = IOFactory::load($file);
    //
    //    foreach ($excel->getAllSheets() as $index => $sheet) {
    //        $parse = ParseSheet::parse($sheet, [
    //            'index' => [
    //                'is' => function ($value) {
    //                    return '序号' === $value;
    //                },
    //                'format' => function ($value) {
    //                    return sprintf('系列号: %s', $value);
    //                },
    //            ],
    //            'name' => [
    //                'is' => function ($value) {
    //                    return '姓名' === $value;
    //                },
    //            ],
    //        ]);
    //
    //        $parse->eachWithCheck(function ($fields, $index) {
    //            return ['index' => 'error'];
    //        });
    //
    //        $parse->dumpErrors();
    //    }
    //
    //    /** 保存错误文件 */
    //    $writer = new Xlsx($excel);
    //    $writer->save(sprintf('%s.error.%s.xlsx', $file, __FUNCTION__));
    //}

    /**
     * @throws Exception
     */
    public function testHandle1()
    {
        $file = __DIR__.'/resources/test.xlsx';
        $excel = IOFactory::load($file);

        foreach ($excel->getAllSheets() as $index => $sheet) {
            $parse = ParseSheet::parse($sheet, [
                'index' => [
                    'is' => function ($value) {
                        return '序号' === $value;
                    },
                    'format' => function ($value) {
                        return sprintf('系列号: %s', $value);
                    },
                ],

                'id_code' => [
                    'is' => function ($value) {
                        return '身份证号' === $value;
                    },
                    'format' => function ($value) {
                        return trim(strtoupper($value), '\'');
                    },
                ],

                'work' => [
                    'is' => function ($value) {
                        return 0 < preg_match("/岗位名称\s+（工种）/", $value);
                    },
                    'format' => function ($value) {
                        return $value;
                    },
                ],

                'code' => [
                    'is' => function ($value) {
                        return '原证书编号' === $value;
                    },
                ],
            ]);

            $parse->eachWithCheck(function ($fields, $index) {
                return [];
            });

            $parse->dumpErrors();

            /** 保存错误文件 */
            $writer = new Xlsx($excel);
            $writer->save(sprintf('%s.error.%s.xlsx', $file, __FUNCTION__));

            $this->assertTrue(true);
        }
    }

    /**
     * 回归测试: 数字单元格的 numberFormat 不能把原始数字截断
     *
     * 历史 bug: getFormattedValue 会按 cell 的 numberFormat 把数字渲染成
     * "1,000" / "¥1,000" / "50%" 等字符串, 调用方 floatval 时会从第一个非数字
     * 字符截断 → 1000 变 1, 百分比 50% 变 0 等等.
     *
     * 修复后: 数字 / 公式且非日期单元格应直接返回原始数字; 日期仍返回字符串;
     * 字符串 / 布尔 / 空值与之前一致.
     *
     * @see docs/fix-numeric-format-truncation.md
     */
    public function testNumericFormatNotTruncated()
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

        /** 千分位: 之前会被渲染成 "1,000" */
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
            ->addHeaderPattern('amount', function ($v) { return '金额' === $v; })
            ->addHeaderPattern('currency', function ($v) { return '货币' === $v; })
            ->addHeaderPattern('percent', function ($v) { return '百分比' === $v; })
            ->addHeaderPattern('date', function ($v) { return '日期' === $v; })
            ->addHeaderPattern('formula', function ($v) { return '公式' === $v; })
            ->addHeaderPattern('plain', function ($v) { return '普通数字' === $v; })
            ->addHeaderPattern('text', function ($v) { return '字符串' === $v; });

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
