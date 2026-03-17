<?php

namespace HughCube\Spreadsheet\Exporter;

class Coordinate
{
    /**
     * 列字母转 0-based 索引, 如 'A' -> 0, 'Z' -> 25, 'AA' -> 26
     */
    public static function columnLetterToIndex(string $letter): int
    {
        $letter = strtoupper($letter);
        $index = 0;
        for ($i = 0; $i < strlen($letter); $i++) {
            $index = $index * 26 + (ord($letter[$i]) - ord('A') + 1);
        }
        return $index - 1;
    }

    /**
     * 0-based 索引转列字母, 如 0 -> 'A', 25 -> 'Z', 26 -> 'AA'
     */
    public static function indexToColumnLetter(int $index): string
    {
        $letter = '';
        $index++;
        while ($index > 0) {
            $index--;
            $letter = chr(ord('A') + ($index % 26)) . $letter;
            $index = intdiv($index, 26);
        }
        return $letter;
    }

    /**
     * 解析单元格地址, 如 'A1' -> [0, 0], 'B3' -> [2, 1] (row, col) 均为 0-based
     *
     * @return array{0: int, 1: int} [row, col]
     */
    public static function parseCellAddress(string $address): array
    {
        preg_match('/^([A-Za-z]+)(\d+)$/', $address, $m);
        return [(int)$m[2] - 1, self::columnLetterToIndex($m[1])];
    }

    /**
     * 解析范围, 如 'A1:C3' -> [0, 0, 2, 2] (startRow, startCol, endRow, endCol) 均为 0-based
     *
     * @return array{0: int, 1: int, 2: int, 3: int}
     */
    public static function parseRange(string $range): array
    {
        [$start, $end] = explode(':', $range);
        [$sr, $sc] = self::parseCellAddress($start);
        [$er, $ec] = self::parseCellAddress($end);
        return [$sr, $sc, $er, $ec];
    }
}
