<?php

namespace HughCube\Spreadsheet\Tests;

use HughCube\Spreadsheet\Models\Header;

class HeaderTest extends TestCase
{
    public function testConstruct()
    {
        $header = new Header('A', '序号');
        $this->assertSame('A', $header->getIndex());
        $this->assertSame('序号', $header->getTitle());
    }

    public function testFormatValueWithoutFormatter()
    {
        $header = new Header('B', '姓名');
        $this->assertSame('张三', $header->formatValue('张三'));
        $this->assertNull($header->formatValue(null));
        $this->assertSame('', $header->formatValue(''));
        $this->assertSame(123, $header->formatValue(123));
    }

    public function testFormatValueWithFormatter()
    {
        $header = new Header('C', '金额', function ($value) {
            return round((float) $value, 2);
        });
        $this->assertSame(3.14, $header->formatValue('3.1415'));
        $this->assertSame(0.0, $header->formatValue(null));
    }

    public function testFormatValueWithTrimFormatter()
    {
        $header = new Header('D', '身份证号', function ($value) {
            return trim(strtoupper((string) $value), '\'');
        });
        $this->assertSame('320219X', $header->formatValue("'320219x'"));
    }
}
