<?php

namespace HughCube\Spreadsheet\Exporter;

class Style
{
    protected ?bool $bold = null;
    protected ?bool $italic = null;
    protected ?bool $underline = null;
    protected ?bool $strikethrough = null;
    protected ?float $fontSize = null;
    protected ?string $fontFamily = null;
    protected ?string $fontColor = null;
    protected ?string $bgColor = null;
    protected ?string $horizontalAlign = null;
    protected ?string $verticalAlign = null;
    protected ?bool $wrapText = null;
    protected ?string $numberFormat = null;

    public static function make(): static
    {
        /** @phpstan-ignore new.static */
        return new static();
    }

    public function bold(?bool $value = true): static
    {
        $this->bold = $value;
        return $this;
    }

    public function italic(?bool $value = true): static
    {
        $this->italic = $value;
        return $this;
    }

    public function underline(?bool $value = true): static
    {
        $this->underline = $value;
        return $this;
    }

    public function strikethrough(?bool $value = true): static
    {
        $this->strikethrough = $value;
        return $this;
    }

    public function fontSize(?float $value): static
    {
        $this->fontSize = $value;
        return $this;
    }

    public function fontFamily(?string $value): static
    {
        $this->fontFamily = $value;
        return $this;
    }

    public function fontColor(?string $value): static
    {
        $this->fontColor = $value;
        return $this;
    }

    public function bgColor(?string $value): static
    {
        $this->bgColor = $value;
        return $this;
    }

    public function horizontalAlign(?string $value): static
    {
        $this->horizontalAlign = $value;
        return $this;
    }

    public function verticalAlign(?string $value): static
    {
        $this->verticalAlign = $value;
        return $this;
    }

    public function wrapText(?bool $value = true): static
    {
        $this->wrapText = $value;
        return $this;
    }

    public function numberFormat(?string $value): static
    {
        $this->numberFormat = $value;
        return $this;
    }

    public function getBold(): ?bool
    {
        return $this->bold;
    }

    public function getItalic(): ?bool
    {
        return $this->italic;
    }

    public function getUnderline(): ?bool
    {
        return $this->underline;
    }

    public function getStrikethrough(): ?bool
    {
        return $this->strikethrough;
    }

    public function getFontSize(): ?float
    {
        return $this->fontSize;
    }

    public function getFontFamily(): ?string
    {
        return $this->fontFamily;
    }

    public function getFontColor(): ?string
    {
        return $this->fontColor;
    }

    public function getBgColor(): ?string
    {
        return $this->bgColor;
    }

    public function getHorizontalAlign(): ?string
    {
        return $this->horizontalAlign;
    }

    public function getVerticalAlign(): ?string
    {
        return $this->verticalAlign;
    }

    public function getWrapText(): ?bool
    {
        return $this->wrapText;
    }

    public function getNumberFormat(): ?string
    {
        return $this->numberFormat;
    }

    /**
     * 合并样式, $other 中非 null 的属性覆盖当前值
     */
    public function merge(?Style $other): static
    {
        if (null === $other) {
            return clone $this;
        }

        $merged = clone $this;
        foreach (get_object_vars($other) as $prop => $value) {
            if ($value !== null) {
                $merged->$prop = $value;
            }
        }
        return $merged;
    }

    public function isEmpty(): bool
    {
        foreach (get_object_vars($this) as $value) {
            if ($value !== null) {
                return false;
            }
        }
        return true;
    }

    /**
     * 转为数组, 用于序列化/缓存 key 等外部场景
     *
     * @return array<string, mixed>
     */
    public function toArray(): array
    {
        return get_object_vars($this);
    }
}
