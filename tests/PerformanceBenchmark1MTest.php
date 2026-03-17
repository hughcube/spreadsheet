<?php

namespace HughCube\Spreadsheet\Tests;

use Generator;
use HughCube\Spreadsheet\Exporter\Style;
use HughCube\Spreadsheet\SheetExporter;

class PerformanceBenchmark1MTest extends TestCase
{
    protected function getTempFile(string $ext = 'xlsx'): string
    {
        $file = sys_get_temp_dir() . '/perf_1m_' . uniqid() . '.' . $ext;
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

    protected function dataGenerator(int $rowCount): Generator
    {
        for ($i = 0; $i < $rowCount; $i++) {
            yield ["name_{$i}", $i, "city_{$i}", "email_{$i}@test.com", $i * 1.5];
        }
    }

    protected function runBenchmark(string $driverName, string $ext, int $rowCount, bool $withStyle): void
    {
        $file = $this->getTempFile($ext);

        $memBefore = memory_get_usage(true);
        $timeStart = microtime(true);

        $exporter = SheetExporter::create($driverName);
        $exporter->openFile($file);
        $exporter->addSheet('Sheet1');

        if ($withStyle) {
            $exporter->setDefaultStyle(Style::make()->fontSize(11.0)->fontFamily('Arial'));
            $exporter->setColumnStyle('A', Style::make()->bold());
            $exporter->setColumnStyle('E', Style::make()->numberFormat('#,##0.00'));
            $exporter->setColumnWidth('A', 20.0);
            $exporter->setColumnWidth('B', 10.0);
            $exporter->writeRow(
                ['Name', 'ID', 'City', 'Email', 'Score'],
                Style::make()->bold()->bgColor('#4472C4')->fontColor('#FFFFFF')
            );
        }

        $exporter->writeRows($this->dataGenerator($rowCount));
        $exporter->close();

        $timeEnd = microtime(true);
        $memPeak = memory_get_peak_usage(true);
        $memUsed = $memPeak - $memBefore;
        $duration = $timeEnd - $timeStart;
        $fileSize = filesize($file);

        $label = $withStyle ? '性能+样式' : '性能';
        echo sprintf(
            "\n[%s] %s | %s行 | 耗时: %.2fs | 内存增量: %.1fMB | 峰值: %.1fMB | 文件: %.1fMB\n",
            $label,
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

    public function testCsv1MNoStyle(): void
    {
        $this->runBenchmark('csv', 'csv', 1_000_000, false);
    }

    public function testCsv1MWithStyle(): void
    {
        $this->runBenchmark('csv', 'csv', 1_000_000, true);
    }

    public function testOpenSpout1MNoStyle(): void
    {
        $this->runBenchmark('openspout', 'xlsx', 1_000_000, false);
    }

    public function testOpenSpout1MWithStyle(): void
    {
        $this->runBenchmark('openspout', 'xlsx', 1_000_000, true);
    }

    public function testXlswriter1MNoStyle(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }
        $this->runBenchmark('xlswriter', 'xlsx', 1_000_000, false);
    }

    public function testXlswriter1MWithStyle(): void
    {
        if (!extension_loaded('xlswriter')) {
            $this->markTestSkipped('xlswriter 扩展未安装');
        }
        $this->runBenchmark('xlswriter', 'xlsx', 1_000_000, true);
    }
}
