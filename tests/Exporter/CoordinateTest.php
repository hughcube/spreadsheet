<?php

namespace HughCube\Spreadsheet\Tests\Exporter;

use HughCube\Spreadsheet\Exporter\Coordinate;
use HughCube\Spreadsheet\Tests\TestCase;

class CoordinateTest extends TestCase
{
    public function testColumnLetterToIndex(): void
    {
        $this->assertSame(0, Coordinate::columnLetterToIndex('A'));
        $this->assertSame(1, Coordinate::columnLetterToIndex('B'));
        $this->assertSame(25, Coordinate::columnLetterToIndex('Z'));
        $this->assertSame(26, Coordinate::columnLetterToIndex('AA'));
        $this->assertSame(27, Coordinate::columnLetterToIndex('AB'));
        $this->assertSame(701, Coordinate::columnLetterToIndex('ZZ'));
        $this->assertSame(702, Coordinate::columnLetterToIndex('AAA'));
    }

    public function testColumnLetterToIndexCaseInsensitive(): void
    {
        $this->assertSame(0, Coordinate::columnLetterToIndex('a'));
        $this->assertSame(26, Coordinate::columnLetterToIndex('aa'));
    }

    public function testIndexToColumnLetter(): void
    {
        $this->assertSame('A', Coordinate::indexToColumnLetter(0));
        $this->assertSame('B', Coordinate::indexToColumnLetter(1));
        $this->assertSame('Z', Coordinate::indexToColumnLetter(25));
        $this->assertSame('AA', Coordinate::indexToColumnLetter(26));
        $this->assertSame('AB', Coordinate::indexToColumnLetter(27));
        $this->assertSame('ZZ', Coordinate::indexToColumnLetter(701));
        $this->assertSame('AAA', Coordinate::indexToColumnLetter(702));
    }

    public function testRoundTrip(): void
    {
        for ($i = 0; $i < 1000; $i++) {
            $letter = Coordinate::indexToColumnLetter($i);
            $this->assertSame($i, Coordinate::columnLetterToIndex($letter));
        }
    }

    public function testParseCellAddress(): void
    {
        $this->assertSame([0, 0], Coordinate::parseCellAddress('A1'));
        $this->assertSame([2, 1], Coordinate::parseCellAddress('B3'));
        $this->assertSame([99, 26], Coordinate::parseCellAddress('AA100'));
    }

    public function testParseRange(): void
    {
        $this->assertSame([0, 0, 2, 2], Coordinate::parseRange('A1:C3'));
        $this->assertSame([0, 0, 0, 3], Coordinate::parseRange('A1:D1'));
        $this->assertSame([4, 2, 9, 5], Coordinate::parseRange('C5:F10'));
    }
}
