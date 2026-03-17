<?php

namespace HughCube\Spreadsheet\Tests;

use Generator;
use HughCube\Spreadsheet\Exporter\Drivers\CsvDriver;
use HughCube\Spreadsheet\Exporter\Drivers\OpenSpoutDriver;
use HughCube\Spreadsheet\Exporter\Drivers\PhpSpreadsheetDriver;
use HughCube\Spreadsheet\Exporter\Drivers\XlswriterDriver;
use HughCube\Spreadsheet\Exporter\Style;
use HughCube\Spreadsheet\SheetExporter;
use LogicException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PHPUnit\Framework\Attributes\DataProvider;

class SheetExporterTest extends TestCase
{
    protected function getTempFile(string $ext = 'xlsx'): string
    {
        $file = sys_get_temp_dir() . '/sheet_exporter_test_' . uniqid() . '.' . $ext;
        $this->tempFiles[] = $file;
        return $file;
    }

    /** @var string[] */
    protected array $tempFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->tempFiles as $file) {
            if (file_exists($file)) {
                @unlink($file);
            }
        }
        parent::tearDown();
    }

    // ==================== 工厂和驱动检测 ====================

    public function testCreateWithDriverName(): void
    {
        $this->assertInstanceOf(CsvDriver::class, SheetExporter::create('csv')->getDriver());
        $this->assertInstanceOf(PhpSpreadsheetDriver::class, SheetExporter::create('phpspreadsheet')->getDriver());
        $this->assertInstanceOf(OpenSpoutDriver::class, SheetExporter::create('openspout')->getDriver());
        $this->assertInstanceOf(OpenSpoutDriver::class, SheetExporter::create('spout')->getDriver());
    }

    public function testCreateWithXlswriter(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }
        $this->assertInstanceOf(XlswriterDriver::class, SheetExporter::create('xlswriter')->getDriver());
    }

    public function testCreateWithDriverInstance(): void
    {
        $driver = new CsvDriver();
        $exporter = SheetExporter::create($driver);
        $this->assertSame($driver, $exporter->getDriver());
    }

    public function testCreateWithNull(): void
    {
        $exporter = SheetExporter::create();
        $this->assertNotNull($exporter->getDriver());
    }

    public function testCreateWithAuto(): void
    {
        $exporter = SheetExporter::create('auto');
        $this->assertNotNull($exporter->getDriver());
    }

    public function testCreateWithUnknownDriverThrows(): void
    {
        $this->expectException(\InvalidArgumentException::class);
        SheetExporter::create('unknown');
    }

    // ==================== 生命周期 ====================

    public function testWriteRowWithoutOpenThrows(): void
    {
        $exporter = SheetExporter::create('csv');
        $this->expectException(LogicException::class);
        $exporter->writeRow(['a']);
    }

    public function testOpenFileTwiceThrows(): void
    {
        $file = $this->getTempFile('csv');
        $exporter = SheetExporter::create('csv');
        $exporter->openFile($file);
        $this->expectException(LogicException::class);
        $exporter->openFile($file);
    }

    public function testCloseWhenNotOpenIsNoop(): void
    {
        $exporter = SheetExporter::create('csv');
        $exporter->close(); // 不应抛异常
        $this->assertTrue(true);
    }

    // ==================== CSV 驱动功能测试 ====================

    public function testCsvBasicExport(): void
    {
        $file = $this->getTempFile('csv');

        $exporter = SheetExporter::create('csv');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['姓名', '年龄', '城市']);
        $exporter->writeRow(['张三', 25, '北京']);
        $exporter->writeRow(['李四', 30, '上海']);
        $exporter->close();

        $this->assertFileExists($file);
        $content = file_get_contents($file);

        /** 检查 BOM */
        $this->assertStringStartsWith("\xEF\xBB\xBF", $content);

        $lines = explode("\n", trim(substr($content, 3)));
        $this->assertCount(3, $lines);
        $this->assertStringContainsString('张三', $lines[1]);
    }

    public function testCsvAppendRow(): void
    {
        $file = $this->getTempFile('csv');

        $exporter = SheetExporter::create('csv');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['header1', 'header2']);
        $exporter->appendRow(['val1', 'val2']);
        $exporter->appendRows([['val3', 'val4'], ['val5', 'val6']]);
        $exporter->close();

        $content = file_get_contents($file);
        $lines = explode("\n", trim(substr($content, 3)));
        $this->assertCount(4, $lines);
    }

    public function testCsvWriteRows(): void
    {
        $file = $this->getTempFile('csv');

        $exporter = SheetExporter::create('csv');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRows([['a', 'b'], ['c', 'd'], ['e', 'f']]);
        $exporter->close();

        $content = file_get_contents($file);
        $lines = explode("\n", trim(substr($content, 3)));
        $this->assertCount(3, $lines);
    }

    public function testCsvWriteRowsWithGenerator(): void
    {
        $file = $this->getTempFile('csv');

        $generator = function (): Generator {
            for ($i = 0; $i < 100; $i++) {
                yield ["row_{$i}", $i];
            }
        };

        $exporter = SheetExporter::create('csv');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRows($generator());
        $exporter->close();

        $content = file_get_contents($file);
        $lines = explode("\n", trim(substr($content, 3)));
        $this->assertCount(100, $lines);
    }

    public function testCsvSupportsStyleReturnsFalse(): void
    {
        $this->assertFalse((new CsvDriver())->supportsStyle());
    }

    // ==================== PhpSpreadsheet 驱动功能测试 ====================

    public function testPhpSpreadsheetBasicExport(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('数据');
        $exporter->writeRow(['姓名', '年龄']);
        $exporter->writeRow(['张三', 25]);
        $exporter->writeRow(['李四', 30]);
        $exporter->close();

        $this->assertFileExists($file);

        /** 回读验证 */
        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertSame('数据', $sheet->getTitle());
        $this->assertSame('姓名', $sheet->getCell('A1')->getValue());
        $this->assertSame('张三', $sheet->getCell('A2')->getValue());
        $this->assertSame(25, $sheet->getCell('B2')->getValue());
        $this->assertSame(30, $sheet->getCell('B3')->getValue());
        $spreadsheet->disconnectWorksheets();
    }

    public function testPhpSpreadsheetStyleHierarchy(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        /** 表默认: 字号11 */
        $exporter->setDefaultStyle(Style::make()->fontSize(11.0));

        /** 列 A: 加粗 */
        $exporter->setColumnStyle('A', Style::make()->bold());

        /** 第 1 行(内联): 背景灰色 */
        $exporter->writeRow(['Name', 'Age'], Style::make()->bgColor('#CCCCCC'));

        /** 预设单元格 B2: 红色字 */
        $exporter->setCellStyle('B2', Style::make()->fontColor('#FF0000'));
        $exporter->writeRow(['John', 25]);

        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();

        /** A1: 表默认(11) + 列A(bold) + 行1(bgColor) */
        $styleA1 = $sheet->getStyle('A1');
        $this->assertTrue($styleA1->getFont()->getBold());
        $this->assertSame(11.0, $styleA1->getFont()->getSize());
        $this->assertSame('CCCCCC', $styleA1->getFill()->getStartColor()->getRGB());

        /** B1: 表默认(11) + 行1(bgColor), 无 bold */
        $styleB1 = $sheet->getStyle('B1');
        $this->assertFalse($styleB1->getFont()->getBold());
        $this->assertSame('CCCCCC', $styleB1->getFill()->getStartColor()->getRGB());

        /** A2: 表默认(11) + 列A(bold), 无 bgColor */
        $styleA2 = $sheet->getStyle('A2');
        $this->assertTrue($styleA2->getFont()->getBold());

        /** B2: 表默认(11) + 单元格(fontColor: red) */
        $styleB2 = $sheet->getStyle('B2');
        $this->assertSame('FF0000', $styleB2->getFont()->getColor()->getRGB());

        $spreadsheet->disconnectWorksheets();
    }

    public function testPhpSpreadsheetColumnWidth(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->setColumnWidth('A', 30.0);
        $exporter->setColumnWidth('B', 15.0);
        $exporter->writeRow(['Name', 'Age']);
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertEqualsWithDelta(30.0, $sheet->getColumnDimension('A')->getWidth(), 0.1);
        $this->assertEqualsWithDelta(15.0, $sheet->getColumnDimension('B')->getWidth(), 0.1);
        $spreadsheet->disconnectWorksheets();
    }

    public function testPhpSpreadsheetRowHeight(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->setRowHeight(1, 40.0);
        $exporter->writeRow(['Name', 'Age']);
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertEqualsWithDelta(40.0, $sheet->getRowDimension(1)->getRowHeight(), 0.1);
        $spreadsheet->disconnectWorksheets();
    }

    public function testPhpSpreadsheetMergeCells(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['合并标题', '', '']);
        $exporter->writeRow(['A', 'B', 'C']);
        $exporter->mergeCells('A1:C1');
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $merges = $sheet->getMergeCells();
        $this->assertNotEmpty($merges);
        $this->assertContains('A1:C1', $merges);
        $spreadsheet->disconnectWorksheets();
    }

    public function testPhpSpreadsheetMultiSheet(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);

        $exporter->addSheet('用户');
        $exporter->writeRow(['姓名', '年龄']);
        $exporter->writeRow(['张三', 25]);

        $exporter->addSheet('订单');
        $exporter->writeRow(['订单号', '金额']);
        $exporter->writeRow(['ORD001', 99.99]);

        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $this->assertSame(2, $spreadsheet->getSheetCount());
        $this->assertSame('用户', $spreadsheet->getSheet(0)->getTitle());
        $this->assertSame('订单', $spreadsheet->getSheet(1)->getTitle());
        $this->assertSame('张三', $spreadsheet->getSheet(0)->getCell('A2')->getValue());
        $this->assertSame('ORD001', $spreadsheet->getSheet(1)->getCell('A2')->getValue());
        $spreadsheet->disconnectWorksheets();
    }

    public function testPhpSpreadsheetTapNativeInstance(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['Name', 'Age']);

        /** 通过 tap 访问原生实例 */
        $tapped = false;
        $exporter->tap(function ($native) use (&$tapped) {
            $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Spreadsheet::class, $native);
            $tapped = true;
        });
        $this->assertTrue($tapped);

        $exporter->close();
    }

    public function testPhpSpreadsheetBeforeCloseCallback(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['Name', 'Age']);

        $called = false;
        $exporter->close(function ($native) use (&$called) {
            $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Spreadsheet::class, $native);
            $called = true;
        });
        $this->assertTrue($called);
    }

    public function testPhpSpreadsheetPresetRowStyle(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        /** 预设第 1 行样式(1-based) */
        $exporter->setRowStyle(1, Style::make()->bold());
        $exporter->writeRow(['Name', 'Age']);
        $exporter->writeRow(['John', 25]);
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertTrue($sheet->getStyle('A1')->getFont()->getBold());
        $this->assertFalse($sheet->getStyle('A2')->getFont()->getBold());
        $spreadsheet->disconnectWorksheets();
    }

    public function testPhpSpreadsheetAllStyleProperties(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        $style = Style::make()
            ->bold()
            ->italic()
            ->underline()
            ->strikethrough()
            ->fontSize(16.0)
            ->fontFamily('微软雅黑')
            ->fontColor('#0000FF')
            ->bgColor('#FFFF00')
            ->horizontalAlign('center')
            ->verticalAlign('top')
            ->wrapText()
            ->numberFormat('#,##0.00');

        $exporter->writeRow([12345.678], $style);
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $cellStyle = $sheet->getStyle('A1');

        $font = $cellStyle->getFont();
        $this->assertTrue($font->getBold());
        $this->assertTrue($font->getItalic());
        $this->assertTrue($font->getStrikethrough());
        $this->assertSame(16.0, $font->getSize());
        $this->assertSame('微软雅黑', $font->getName());
        $this->assertSame('0000FF', $font->getColor()->getRGB());

        $this->assertSame('FFFF00', $cellStyle->getFill()->getStartColor()->getRGB());
        $this->assertTrue($cellStyle->getAlignment()->getWrapText());
        $this->assertSame('#,##0.00', $cellStyle->getNumberFormat()->getFormatCode());

        $spreadsheet->disconnectWorksheets();
    }

    // ==================== OpenSpout 驱动功能测试 ====================

    public function testOpenSpoutBasicExport(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('openspout');
        $exporter->openFile($file);
        $exporter->addSheet('数据');
        $exporter->writeRow(['姓名', '年龄']);
        $exporter->writeRow(['张三', 25]);
        $exporter->close();

        $this->assertFileExists($file);

        /** 使用 PhpSpreadsheet 回读 OpenSpout 生成的文件 */
        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertSame('数据', $sheet->getTitle());
        $this->assertSame('姓名', $sheet->getCell('A1')->getFormattedValue());
        $this->assertSame('张三', $sheet->getCell('A2')->getFormattedValue());
        $spreadsheet->disconnectWorksheets();
    }

    public function testOpenSpoutWithStyle(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('openspout');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['Header'], Style::make()->bold()->fontSize(14.0));
        $exporter->writeRow(['Data']);
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertTrue($sheet->getStyle('A1')->getFont()->getBold());
        $this->assertSame(14.0, $sheet->getStyle('A1')->getFont()->getSize());
        $this->assertFalse($sheet->getStyle('A2')->getFont()->getBold());
        $spreadsheet->disconnectWorksheets();
    }

    public function testOpenSpoutMultiSheet(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('openspout');
        $exporter->openFile($file);

        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['A1']);

        $exporter->addSheet('Sheet2');
        $exporter->writeRow(['B1']);

        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $this->assertSame(2, $spreadsheet->getSheetCount());
        $this->assertSame('Sheet1', $spreadsheet->getSheet(0)->getTitle());
        $this->assertSame('Sheet2', $spreadsheet->getSheet(1)->getTitle());
        $spreadsheet->disconnectWorksheets();
    }

    public function testOpenSpoutNativeInstance(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('openspout');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        $exporter->tap(function ($native) {
            $this->assertInstanceOf(\OpenSpout\Writer\XLSX\Writer::class, $native);
        });

        $exporter->writeRow(['test']);
        $exporter->close();
    }

    // ==================== Xlswriter 驱动功能测试 ====================

    public function testXlswriterBasicExport(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('xlswriter');
        $exporter->openFile($file);
        $exporter->addSheet('数据');
        $exporter->writeRow(['姓名', '年龄']);
        $exporter->writeRow(['张三', 25]);
        $exporter->close();

        $this->assertFileExists($file);

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertSame('姓名', $sheet->getCell('A1')->getValue());
        $this->assertSame('张三', $sheet->getCell('A2')->getValue());
        $this->assertSame(25, $sheet->getCell('B2')->getValue());
        $spreadsheet->disconnectWorksheets();
    }

    public function testXlswriterWithStyle(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('xlswriter');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['Header'], Style::make()->bold());
        $exporter->writeRow(['Data']);
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        $this->assertTrue($sheet->getStyle('A1')->getFont()->getBold());
        $spreadsheet->disconnectWorksheets();
    }

    public function testXlswriterMultiSheet(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('xlswriter');
        $exporter->openFile($file);

        $exporter->addSheet('用户');
        $exporter->writeRow(['姓名']);
        $exporter->writeRow(['张三']);

        $exporter->addSheet('订单');
        $exporter->writeRow(['订单号']);
        $exporter->writeRow(['ORD001']);

        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $this->assertSame(2, $spreadsheet->getSheetCount());
        $spreadsheet->disconnectWorksheets();
    }

    public function testXlswriterNativeInstance(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('xlswriter');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        $exporter->tap(function ($native) {
            $this->assertInstanceOf(\Vtiful\Kernel\Excel::class, $native);
        });

        $exporter->writeRow(['test']);
        $exporter->close();
    }

    // ==================== 样式层级优先级测试(跨驱动) ====================

    #[DataProvider('driverProvider')]
    public function testStylePriorityCellOverridesRow(string $driverName): void
    {
        if ('xlswriter' === $driverName && !extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create($driverName);
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        /** 单元格 A1 红色字, 行样式背景灰色 */
        $exporter->setCellStyle('A1', Style::make()->fontColor('#FF0000'));
        $exporter->writeRow(['Test'], Style::make()->bgColor('#CCCCCC'));
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();

        /** 单元格级 fontColor 存在 */
        $this->assertSame('FF0000', $sheet->getStyle('A1')->getFont()->getColor()->getRGB());
        /** 行级 bgColor 也存在 */
        $this->assertSame('CCCCCC', $sheet->getStyle('A1')->getFill()->getStartColor()->getRGB());
        $spreadsheet->disconnectWorksheets();
    }

    public static function driverProvider(): array
    {
        $drivers = [
            'phpspreadsheet' => ['phpspreadsheet'],
            'openspout' => ['openspout'],
        ];

        if (extension_loaded('xlswriter')) {
            $drivers['xlswriter'] = ['xlswriter'];
        }

        return $drivers;
    }

    // ==================== 样式缓存: 不同样式必须产生不同结果 ====================

    #[DataProvider('driverProvider')]
    public function testDifferentStylesProduceDifferentResults(string $driverName): void
    {
        if ('xlswriter' === $driverName && !extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create($driverName);
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        /** 第 1 行: 加粗红字 */
        $exporter->writeRow(['Bold Red'], Style::make()->bold()->fontColor('#FF0000'));
        /** 第 2 行: 斜体蓝字 */
        $exporter->writeRow(['Italic Blue'], Style::make()->italic()->fontColor('#0000FF'));
        /** 第 3 行: 无样式 */
        $exporter->writeRow(['Plain']);

        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();

        /** 第 1 行: 加粗 + 红字 */
        $this->assertTrue($sheet->getStyle('A1')->getFont()->getBold());
        $this->assertSame('FF0000', $sheet->getStyle('A1')->getFont()->getColor()->getRGB());

        /** 第 2 行: 斜体 + 蓝字, 不加粗 */
        $this->assertTrue($sheet->getStyle('A2')->getFont()->getItalic());
        $this->assertSame('0000FF', $sheet->getStyle('A2')->getFont()->getColor()->getRGB());
        $this->assertFalse($sheet->getStyle('A2')->getFont()->getBold());

        /** 第 3 行: 无特殊样式 */
        $this->assertFalse($sheet->getStyle('A3')->getFont()->getBold());
        $this->assertFalse($sheet->getStyle('A3')->getFont()->getItalic());

        $spreadsheet->disconnectWorksheets();
    }

    // ==================== Xlswriter 第一个 sheet 名称 ====================

    public function testXlswriterFirstSheetName(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('xlswriter');
        $exporter->openFile($file);
        $exporter->addSheet('自定义名称');
        $exporter->writeRow(['test']);
        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $this->assertSame('自定义名称', $spreadsheet->getActiveSheet()->getTitle());
        $spreadsheet->disconnectWorksheets();
    }

    // ==================== 合并单元格(跨 sheet 切换时正确应用) ====================

    public function testMergeCellsAppliedOnSheetSwitch(): void
    {
        $file = $this->getTempFile('xlsx');

        $exporter = SheetExporter::create('phpspreadsheet');
        $exporter->openFile($file);

        $exporter->addSheet('Sheet1');
        $exporter->writeRow(['合并', '', '']);
        $exporter->writeRow(['a', 'b', 'c']);
        $exporter->mergeCells('A1:C1');

        $exporter->addSheet('Sheet2');
        $exporter->writeRow(['test']);

        $exporter->close();

        $spreadsheet = IOFactory::load($file);
        $this->assertContains('A1:C1', $spreadsheet->getSheet(0)->getMergeCells());
        $spreadsheet->disconnectWorksheets();
    }

    // ==================== 流式大数据: Generator ====================

    public function testStreamingWithGenerator(): void
    {
        $file = $this->getTempFile('csv');
        $rowCount = 10000;

        $generator = function () use ($rowCount): Generator {
            for ($i = 0; $i < $rowCount; $i++) {
                yield ["name_{$i}", $i, "city_{$i}"];
            }
        };

        $exporter = SheetExporter::create('csv');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRows($generator());
        $exporter->close();

        $content = file_get_contents($file);
        $lines = explode("\n", trim(substr($content, 3)));
        $this->assertCount($rowCount, $lines);
    }

    // ==================== 性能/压力测试 ====================

    #[DataProvider('performanceDriverProvider')]
    public function testPerformanceExport(string $driverName, string $ext, int $rowCount): void
    {
        if ('xlswriter' === $driverName && !extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile($ext);
        $columns = 5;

        /** 行数据生成器, 不占额外内存 */
        $generator = function () use ($rowCount, $columns): Generator {
            for ($i = 0; $i < $rowCount; $i++) {
                $row = [];
                for ($j = 0; $j < $columns; $j++) {
                    $row[] = "data_{$i}_{$j}";
                }
                yield $row;
            }
        };

        $memBefore = memory_get_usage(true);
        $timeStart = microtime(true);

        $exporter = SheetExporter::create($driverName);
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');
        $exporter->writeRows($generator());
        $exporter->close();

        $timeEnd = microtime(true);
        $memPeak = memory_get_peak_usage(true);
        $memUsed = $memPeak - $memBefore;
        $duration = $timeEnd - $timeStart;
        $fileSize = filesize($file);

        echo sprintf(
            "\n[性能] %s | %s行 | 耗时: %.2fs | 内存增量: %.1fMB | 峰值: %.1fMB | 文件: %.1fMB\n",
            strtoupper($driverName),
            number_format($rowCount),
            $duration,
            $memUsed / 1024 / 1024,
            $memPeak / 1024 / 1024,
            $fileSize / 1024 / 1024,
        );

        $this->assertFileExists($file);
        $this->assertGreaterThan(0, $fileSize);
    }

    public static function performanceDriverProvider(): array
    {
        $drivers = [
            'csv_10k' => ['csv', 'csv', 10_000],
            'csv_100k' => ['csv', 'csv', 100_000],
            'openspout_10k' => ['openspout', 'xlsx', 10_000],
            'openspout_100k' => ['openspout', 'xlsx', 100_000],
        ];

        if (extension_loaded('xlswriter')) {
            $drivers['xlswriter_10k'] = ['xlswriter', 'xlsx', 10_000];
            $drivers['xlswriter_100k'] = ['xlswriter', 'xlsx', 100_000];
        }

        /** PhpSpreadsheet 内存占用高, 只测小量 */
        $drivers['phpspreadsheet_10k'] = ['phpspreadsheet', 'xlsx', 10_000];

        return $drivers;
    }

    /**
     * 带样式的性能测试, 验证样式解析不会造成明显性能下降
     */
    #[DataProvider('performanceWithStyleDriverProvider')]
    public function testPerformanceExportWithStyle(string $driverName, string $ext, int $rowCount): void
    {
        if ('xlswriter' === $driverName && !extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }

        $file = $this->getTempFile($ext);

        $generator = function () use ($rowCount): Generator {
            for ($i = 0; $i < $rowCount; $i++) {
                yield ["name_{$i}", $i, "city_{$i}", "email_{$i}@test.com", $i * 1.5];
            }
        };

        $memBefore = memory_get_usage(true);
        $timeStart = microtime(true);

        $exporter = SheetExporter::create($driverName);
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        /** 表级默认样式 */
        $exporter->setDefaultStyle(Style::make()->fontSize(11.0)->fontFamily('Arial'));

        /** 列样式 */
        $exporter->setColumnStyle('A', Style::make()->bold());
        $exporter->setColumnStyle('E', Style::make()->numberFormat('#,##0.00'));

        /** 列宽 */
        $exporter->setColumnWidth('A', 20.0);
        $exporter->setColumnWidth('B', 10.0);

        /** 带行内联样式的表头 */
        $exporter->writeRow(['Name', 'ID', 'City', 'Email', 'Score'], Style::make()->bold()->bgColor('#4472C4')->fontColor('#FFFFFF'));

        /** 数据行 */
        $exporter->writeRows($generator());

        $exporter->close();

        $timeEnd = microtime(true);
        $memPeak = memory_get_peak_usage(true);
        $memUsed = $memPeak - $memBefore;
        $duration = $timeEnd - $timeStart;
        $fileSize = filesize($file);

        echo sprintf(
            "\n[性能+样式] %s | %s行 | 耗时: %.2fs | 内存增量: %.1fMB | 峰值: %.1fMB | 文件: %.1fMB\n",
            strtoupper($driverName),
            number_format($rowCount),
            $duration,
            $memUsed / 1024 / 1024,
            $memPeak / 1024 / 1024,
            $fileSize / 1024 / 1024,
        );

        $this->assertFileExists($file);
        $this->assertGreaterThan(0, $fileSize);
    }

    public static function performanceWithStyleDriverProvider(): array
    {
        $drivers = [
            'csv_style_100k' => ['csv', 'csv', 100_000],
            'openspout_style_10k' => ['openspout', 'xlsx', 10_000],
            'openspout_style_100k' => ['openspout', 'xlsx', 100_000],
        ];

        if (extension_loaded('xlswriter')) {
            $drivers['xlswriter_style_10k'] = ['xlswriter', 'xlsx', 10_000];
            $drivers['xlswriter_style_100k'] = ['xlswriter', 'xlsx', 100_000];
        }

        $drivers['phpspreadsheet_style_10k'] = ['phpspreadsheet', 'xlsx', 10_000];

        return $drivers;
    }

    /**
     * 验证流式写入内存不会随行数线性增长
     */
    public function testMemoryDoesNotGrowLinearlyWithRows(): void
    {
        $file = $this->getTempFile('csv');

        $exporter = SheetExporter::create('csv');
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        /** 写 10000 行后记录内存 */
        for ($i = 0; $i < 10_000; $i++) {
            $exporter->writeRow(["row_{$i}", $i, "data_{$i}"]);
        }
        $memAfter10k = memory_get_usage(true);

        /** 再写 90000 行 */
        for ($i = 10_000; $i < 100_000; $i++) {
            $exporter->writeRow(["row_{$i}", $i, "data_{$i}"]);
        }
        $memAfter100k = memory_get_usage(true);

        $exporter->close();

        /**
         * 内存增量应该极小(几百KB以内), 不应随行数 10x 而 10x 增长
         * 如果旧的缓存方案, 这里内存会增长约 10 倍
         */
        $memGrowth = $memAfter100k - $memAfter10k;
        echo sprintf("\n[内存线性测试] 10K→100K 内存增量: %.2fMB\n", $memGrowth / 1024 / 1024);

        $this->assertLessThan(10 * 1024 * 1024, $memGrowth, '内存增量不应超过 10MB (9倍行数增长)');
    }
}
