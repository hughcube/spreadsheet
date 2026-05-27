# 修复：数字单元格的千分位 / 货币格式被 `floatval` 截断

> 涉及版本：**v1.0.10**, **v1.1.1**, **v2.1.1**, **v3.0.1**

## 一句话总结

`SheetParser` 之前默认对所有单元格调用 `Cell::getFormattedValue()`（v2/v3 是底层走 `Worksheet::rangeToArray(..., formatData=true, ...)`），数字单元格如果带千分位 / 货币 / 百分比等数字格式，会被 `NumberFormat::toFormattedString()` 转成 **显示字符串**（如 `"1,000"`、`"￥1,000.50"`），调用方一旦用 `floatval()` 或 `(int)` 把它转回数字，PHP 在第一个非数字字符处停止解析，导致金额被截断成 `1`、`0` 之类的错误值。

修复后，**数字 / 公式且非日期单元格** 走 `Cell::getCalculatedValue()` 或 `rangeToArray(..., formatData=false, ...)` 拿原始数字；**日期 / 字符串 / 布尔 / 错误等** 仍走 `getFormattedValue()`，保留可读形式。

## 触发场景

### 一个真实事故（金额 1000 被推送成 1）

调用代码大致如下：

```php
SheetParser::parse($sheet)
    ->addHeaderPattern(
        'paid_amount',
        is:     fn($v) => preg_replace("/[\s']/", '', $v) === '支付金额',
        format: fn($v) => floatval($v),
    )
    ->eachWithCheck(function ($fields) {
        // $fields['paid_amount'] 期望是 1000.0，实际拿到 1.0
        thirdPartyApi->push(['payMoney' => $fields['paid_amount']]);
    });
```

Excel 单元格的真实情况：

```xml
<!-- xl/worksheets/sheet1.xml -->
<c r="I2" s="9"><v>1000</v></c>

<!-- xl/styles.xml 中 cellXfs 索引 9 引用 numFmtId=3, 即内置千分位格式 #,##0 -->
```

| 阶段 | 值 | 类型 |
|---|---|---|
| Excel 存储 | `1000` | 数字 |
| 应用 numFmt `#,##0` 后显示 | `"1,000"` | 字符串 |
| `getFormattedValue()` 返回 | `"1,000"` | 字符串 |
| `floatval("1,000")` | `1.0` | float（**被截断**） |

第三方系统因此收到 `payMoney=1`，但实际成交金额是 `1000` —— 直接导致后续退费、对账、报表全错。

### 其他会踩到同一坑的格式

只要数字单元格应用了任何"显示字符串带非数字字符"的 numberFormat，都会有同样问题。常见的：

| Excel numberFormat | 显示效果 | `floatval` 结果 |
|---|---|---|
| `#,##0` | `1,000` | `1` |
| `#,##0.00` | `1,234.56` | `1` |
| `"￥"#,##0` | `￥1,000` | `0` |
| `0%` | `50%`（原始 0.5） | `0` |
| `0.00E+00` | `1.23E+03` | `1.23`（数量级丢失） |

## 根因详细

PhpSpreadsheet `Cell::getFormattedValue()` 实际就是 `NumberFormat::toFormattedString($cell->getCalculatedValue(), $cell->getStyle()->getNumberFormat()->getFormatCode())`，对数字单元格而言它**有损**——把数字变成字符串。

`Worksheet::rangeToArray()` 第 4 个参数 `$formatData` 默认为 `true`，内部对每个 cell 也调用 `getFormattedValue()`，所以批量读取也会触发。

对**日期单元格**而言，`getFormattedValue()` 是必要的：日期在 Excel 底层就是一个 serial number（如 `46167.638`），不格式化拿出来是没法用的。所以这次修复**保留**日期单元格走格式化路径。

## 修复策略对比

| 版本 | 数据读取实现 | 修复方式 |
|---|---|---|
| **v1.0.10 / v1.1.1** | `getRowIterator` + `getCellIterator` 逐单元格 + `getFormattedValue()` | 在循环里加判断：`isNumericLike && !Date::isDateTime($cell)` 走 `getCalculatedValue()` |
| **v2.1.1 / v3.0.1** | `Worksheet::rangeToArray(..., formatData=true, ...)` 批量读取 | 改 `formatData=false`，再用 `getCellIterator` 找出日期单元格、用 `getFormattedValue()` 单独覆盖 |

两种路径最终行为一致：

```
数字 / 公式 + 非日期 → 原始 / 计算后的数字值
日期                 → 格式化字符串（如 "2026-05-25 15:19:00"）
字符串 / 布尔 / 其他 → getFormattedValue（与之前一致）
异常                 → fallback 到 getValue / 上一步结果
```

## 升级指南

| 项目当前 composer 约束 | 建议升级到 |
|---|---|
| `"1.0.*"` | `v1.0.10`（`composer update hughcube/spreadsheet`） |
| `"1.1.*"` | `v1.1.1` |
| `"2.*"` 或 `"2.1.*"` | `v2.1.1` |
| `"^3.0"` / `"3.0.*"` / `"dev-master"` | `v3.0.1` |

升级后**无需改动调用代码**。

## 向后兼容性

可能感受到差异的场景：

1. **之前依赖 `$fields[key]` 是字符串、然后做字符串操作**
   修复后数字单元格返回 `int|float`，而不是 `"1,000"` 这种带千分位的字符串。
   如果原来用 `preg_replace("/[^0-9.]/", '', $v)` 这种宽松清洗的写法，仍然能工作；但若直接判断 `is_string($v)` 会变成 false。

2. **百分比单元格**
   `0%`、`50%` 这种格式的单元格，之前拿到 `"0%"`、`"50%"`，修复后拿到 `0`、`0.5`（PhpSpreadsheet 内部 0~1 表示百分比）。如果之前在 format 闭包里写过 `rtrim($v, '%') / 100` 之类的兼容代码，**应该改成直接使用，不再除以 100**。

3. **公式单元格**
   之前 `getFormattedValue()` 返回的是公式结果的格式化字符串；修复后通过 `getCalculatedValue()` 拿到的是原始结果数字（保留小数精度）。

如果上述场景对你的项目有影响，可以在调用方的 `format:` 闭包里恢复格式化行为，或固定到 patch 之前的版本（`"1.0.9"` / `"1.1.0"` / `"2.1.0"` / `"3.0.0"`）。

## 涉及的代码改动

- `src/SheetParser.php`：`getDataIterator()` 方法内的单元格读取逻辑

不影响公共 API：方法签名、返回值类型、generator 行为均保持不变。

## 验证方式

一个最小复现单测：

```php
$sheet = (new Spreadsheet())->getActiveSheet();
$sheet->setCellValue('A1', '金额');
$sheet->setCellValue('A2', 1000);
$sheet->getStyle('A2')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1); // #,##0

$amount = null;
SheetParser::parse($sheet)
    ->addHeaderPattern('amount', is: fn($v) => $v === '金额', format: fn($v) => floatval($v))
    ->eachWithCheck(function ($fields) use (&$amount) { $amount = $fields['amount']; });

// 修复前: $amount === 1.0
// 修复后: $amount === 1000.0
self::assertSame(1000.0, $amount);
```
