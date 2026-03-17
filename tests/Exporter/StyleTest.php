<?php

namespace HughCube\Spreadsheet\Tests\Exporter;

use HughCube\Spreadsheet\Exporter\Style;
use HughCube\Spreadsheet\Tests\TestCase;

class StyleTest extends TestCase
{
    public function testMakeReturnsNewInstance(): void
    {
        $style = Style::make();
        $this->assertInstanceOf(Style::class, $style);
        $this->assertTrue($style->isEmpty());
    }

    public function testFluentSettersAndGetters(): void
    {
        $style = Style::make()
            ->bold()
            ->italic()
            ->underline()
            ->strikethrough()
            ->fontSize(14.0)
            ->fontFamily('Arial')
            ->fontColor('#FF0000')
            ->bgColor('#00FF00')
            ->horizontalAlign('center')
            ->verticalAlign('top')
            ->wrapText()
            ->numberFormat('#,##0.00');

        $this->assertTrue($style->getBold());
        $this->assertTrue($style->getItalic());
        $this->assertTrue($style->getUnderline());
        $this->assertTrue($style->getStrikethrough());
        $this->assertSame(14.0, $style->getFontSize());
        $this->assertSame('Arial', $style->getFontFamily());
        $this->assertSame('#FF0000', $style->getFontColor());
        $this->assertSame('#00FF00', $style->getBgColor());
        $this->assertSame('center', $style->getHorizontalAlign());
        $this->assertSame('top', $style->getVerticalAlign());
        $this->assertTrue($style->getWrapText());
        $this->assertSame('#,##0.00', $style->getNumberFormat());
        $this->assertFalse($style->isEmpty());
    }

    public function testSetNullClearsProperty(): void
    {
        $style = Style::make()->bold(true)->bold(null);
        $this->assertNull($style->getBold());
        $this->assertTrue($style->isEmpty());
    }

    public function testMergeOverridesNonNull(): void
    {
        $base = Style::make()->fontSize(11.0)->fontFamily('Arial')->bold();
        $override = Style::make()->fontSize(14.0)->italic();

        $merged = $base->merge($override);

        $this->assertSame(14.0, $merged->getFontSize());
        $this->assertSame('Arial', $merged->getFontFamily());
        $this->assertTrue($merged->getBold());
        $this->assertTrue($merged->getItalic());
    }

    public function testMergeDoesNotMutateOriginal(): void
    {
        $base = Style::make()->fontSize(11.0);
        $override = Style::make()->fontSize(14.0);

        $base->merge($override);

        $this->assertSame(11.0, $base->getFontSize());
    }

    public function testMergeWithNull(): void
    {
        $style = Style::make()->bold();
        $merged = $style->merge(null);

        $this->assertTrue($merged->getBold());
        $this->assertNotSame($style, $merged);
    }

    public function testIsEmpty(): void
    {
        $this->assertTrue(Style::make()->isEmpty());
        $this->assertFalse(Style::make()->bold()->isEmpty());
    }

    public function testMergeChainFourLevels(): void
    {
        $default = Style::make()->fontSize(10.0)->fontFamily('宋体');
        $column = Style::make()->bold();
        $row = Style::make()->bgColor('#EEEEEE');
        $cell = Style::make()->fontColor('#FF0000');

        $resolved = $default->merge($column)->merge($row)->merge($cell);

        $this->assertSame(10.0, $resolved->getFontSize());
        $this->assertSame('宋体', $resolved->getFontFamily());
        $this->assertTrue($resolved->getBold());
        $this->assertSame('#EEEEEE', $resolved->getBgColor());
        $this->assertSame('#FF0000', $resolved->getFontColor());
    }
}
