<?php

namespace HughCube\Spreadsheet\Tests;

use HughCube\Spreadsheet\Models\Header;
use HughCube\Spreadsheet\Models\Headers;

class HeadersTest extends TestCase
{
    public function testConstruct()
    {
        $headers = new Headers(2, [
            'name' => new Header('A', '姓名'),
            'age'  => new Header('B', '年龄'),
        ]);

        $this->assertSame(2, $headers->getIndex());
        $this->assertCount(2, $headers->getHeaders());
    }

    public function testGetMaxColumnSingleLetter()
    {
        $headers = new Headers(1, [
            'a' => new Header('A', 'A'),
            'c' => new Header('C', 'C'),
            'b' => new Header('B', 'B'),
        ]);
        $this->assertSame('C', $headers->getMaxColumn());
    }

    public function testGetMaxColumnMultiLetter()
    {
        $headers = new Headers(1, [
            'x' => new Header('Z', 'Z'),
            'y' => new Header('AA', 'AA'),
            'z' => new Header('AB', 'AB'),
        ]);
        // AB > AA > Z, 修复前 max() 字典序比较会错误地返回 Z
        $this->assertSame('AB', $headers->getMaxColumn());
    }

    public function testGetMaxColumnEmpty()
    {
        $headers = new Headers(1, []);
        $this->assertSame('A', $headers->getMaxColumn());
    }

    public function testGetMaxHeaderIndexDeprecated()
    {
        $headers = new Headers(1, [
            'a' => new Header('D', 'D'),
        ]);
        $this->assertSame('D', $headers->getMaxHeaderIndex());
    }
}
