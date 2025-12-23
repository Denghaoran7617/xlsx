# `tealeg/xlsx` 教程：使用 Go 读写 `xlsx` 文件

## 开始使用

### 下载包

要在代码中导入包，请使用以下行：

```go
import "github.com/Denghaoran7617/xlsx"
```

如果您使用 https://blog.golang.org/using-go-modules[Go 模块]，您的 `go.mod` 文件应包含包的 `require` 行，如下所示：

```go
require github.com/Denghaoran7617/xlsx v3.2.0
```

_如果您想知道为什么我使用 `v3.2.0` 版本标签，请阅读下一章。_

### 我需要哪个版本的包来学习本教程？

`xlsx` 包是一个活跃的项目，因此提供过时的信息没有意义。本教程涵盖包的 `3.x` 版本。如果您使用 `2.x` 分支，也可以使用本教程的很多内容。在这种情况下，某些功能会有所不同（例如 `Column` 相关内容），但新功能以及错误修复非常值得升级。

如果您仍在使用包的 `v1.0.5` 版本，我不能保证您能从本教程中获得很多收获。就我个人而言，我最近才停止使用这个旧版本，当时我有了编写本教程的想法。我很快意识到，咬紧牙关升级代码会让我受益匪浅（一个周末的时间花得很值）。因此，如果您有时间，我强烈建议升级到 `xlsx` 包的 `3.x` 分支。

**注意：** 我确实建议至少使用包的 3.2.0 版本，因为在该版本中，您现在可以询问行或单元格的坐标。在之前的版本中，这些信息未在结构体中导出。

### 您有示例文件吗？

是的。我将为本教程使用一个示例文件（创意命名为 `samplefile.xlsx`），它包含两个工作表。第一个工作表名为 `Sample`，包含以下数据：

| Name    | Date      | Number | Formula |
|---------|-----------|--------|---------|
| Alice   | 3/20/2020 | 24315  | 1215.75 € |
| Bob     | 4/17/2020 | 21345  | 1067.25 € |
| Charlie | 2/8/2020  | 32145  | 1607.25 € |

第二个工作表名为 `Cities`，包含世界上最大城市的列表：

| City        | Country   | Pop(Mio) |
|-------------|-----------|----------|
| Tokyo       | Japan     | 37.40 |
| Dehli       | India     | 28.51 |
| Shanghai    | China     | 25.58 |
| São Paulo   | Brazil    | 21.65 |
| Mexico City | Mexico    | 21.58 |
| Cairo       | Egypt     | 20.07 |
| Mumbai      | India     | 19.98 |

## 打开和创建文件

在创建新文件并需要填充数据之前，让我们先打开现有的示例文件。要打开文件，请使用 `OpenFile()` 函数。

```go
// 打开现有文件
wb, err := xlsx.OpenFile("../samplefile.xlsx")
if err != nil {
    panic(err)
}
// wb 现在包含对工作簿的引用
// 显示工作簿中的所有工作表
fmt.Println("Sheets in this file:")
for i, sh := range wb.Sheets {
    fmt.Println(i, sh.Name)
}
fmt.Println("----")
```

要创建新的空 xlsx 文件，请使用 `NewFile()` 函数。

```go
wb := xlsx.NewFile()
```

此函数返回一个新的 `xlsx.File` 结构体。

## 使用工作表

### 访问工作表

`xlsx.File` 结构体包含一个 `Sheets` 字段，它是工作簿中工作表的指针切片（`[]*xlsx.Sheet`）。您可以使用此字段访问文件中的工作表。

```go
// wb 包含对已打开工作簿的引用
fmt.Println("Workbook contains", len(wb.Sheets), "sheets.")
```

但是，大多数时候您可能希望直接访问特定工作表。为此，请使用 `Sheet` 字段，它是一个以字符串为键、以工作表指针为值的映射（`map[string]*xlsx.Sheet`）。键是工作表的名称。

在我们的示例文件中获取名为 "_Sample_" 的工作表引用的简单方法如下：

```go
sheetName := "Sample"
sh, ok := wb.Sheet[sheetName]
if !ok {
    fmt.Println("Sheet does not exist")
    return
}
fmt.Println("Max row in sheet:", sh.MaxRow)
```

始终确保检查从映射返回的工作表是否存在。否则您会遇到运行时错误，因为在我们示例中 `sh` 的值仍然是 `nil` 值。

### 创建工作表

有两种方法可以向工作簿添加新内容：添加（创建）新工作表或将现有工作表结构体追加到工作簿。让我们从第一种方法开始：

```go
filename := "samplefile.xlsx"
wb, err := xlsx.OpenFile(filename)
if err != nil {
    panic(err)
}
sh, err := wb.AddSheet("My New Sheet")
fmt.Println(err)
fmt.Println(sh)
```

**重要提示：** 添加新工作表时检查错误很重要。我作为经验丰富的错误制造者写下这一点 ;-) – 很容易忘记 Excel 中工作表名称的一些限制。

以下是命名工作表时必须记住的限制：

* 工作表名称的最小长度为 1 个字符。
* 工作表名称的最大长度为 31 个字符。
* 这些特殊字符也不允许：: \ / ? * [ ]

如果违反任何这些规则，`AddSheet()` 函数将返回错误。

第二种方法使用您创建的现有 `xlsx.Sheet` 结构体并调用 `AppendSheet()` 函数：

```go
sh, err := wb.AppendSheet(newSheet, "A new sheet")
```

第一个参数（示例代码行中的 `newSheet`）是包含工作表结构体的变量。第二个参数（`"A new sheet"`）是新工作表的名称。上述命名规则适用。此函数返回指向新追加工作表的指针和错误代码。如果您不需要指针而只想检查错误，可以使用通常的下划线忽略该值。

### 关闭工作表

完成工作表操作并保存工作后，建议在工作表上调用 `Close()`。根据代码中 Geoff 的建议："_移除工作表的依赖资源 - 如果您已完成工作表操作，应调用此函数以清除工作表的持久缓存。通常这发生在您保存更改之后。_"

## 使用行和单元格

### 行

`xlsx.Row` 结构体表示工作表中的单行。您可以使用 `Row(index int)` 函数获取对特定行的引用，该函数返回指向单元格行的指针和错误代码。让我们读取索引为 1 的行（_所有行和列的数字值都是从 0 开始的，因此我们将读取工作表中的第二行_）。

```go
// sh 是对工作表的引用，见上文
row, err := sh.Row(1)
if err != nil {
    panic(err)
}
// 让我们对行做一些操作...
fmt.Println(row)
```

行结构体仅导出两个字段：`Hidden`（一个布尔值，显示行是否隐藏）和 `Sheet`（指向包含该行的工作表的指针）。那么如何访问行中的任何内容呢？我们将在关于单元格的章节中看到，但让我们先看看如何添加和删除行。

#### 我的数据在哪里结束？

很好的问题。我们的示例文件在 `Sample` 工作表中仅包含四行。

| Name    | Date      | Number | Formula |
|---------|-----------|--------|---------|
| Alice   | 3/20/2020 | 24315  | 1215.75 € |
| Bob     | 4/17/2020 | 21345  | 1067.25 € |
| Charlie | 2/8/2020  | 32145  | 1607.25 € |

如果我们尝试检索第 123 行会怎样？好吧，我们不会收到错误，并且会得到一个空行。这就是 `Sheet.MaxRow` 发挥作用的地方。正如您在工作表访问章节中学到的，此字段保存工作表中的行数。

```go
sheetName := "Sample"
sh, ok := wb.Sheet[sheetName]
if !ok {
    fmt.Println("Sheet does not exist")
    return
}
fmt.Println("Max row in sheet:", sh.MaxRow)
```

使用示例文件，输出将是：`Max row in sheet: 4`。*注意*：此值不是从 0 开始的（否则它必须是 3）！当您需要知道工作表中包含数据的行数时，请确保检查 `MaxRow` 的值。

#### 添加行

要在当前数据的末尾添加行，请调用 `Sheet` 的 `AddRow()` 函数。这将返回指向行结构体的指针（`*xlsx.Row`）。不需要错误代码，因为代码只是在数据末尾追加一行（如果需要，添加空行）。

您还可以使用工作表提供的 `AddRowAtIndex(index int)` 函数在工作表中的特定索引位置添加行。此函数返回指向行结构体的指针*并且确实返回错误代码*。此函数还检查索引是否小于 0（因为行索引从 0 开始）或行索引是否大于 `MaxRow`。尝试为上面的示例工作表调用 `row, err := sh.AddRowAtIndex(123)` 将导致 `err` 中的错误和 `row` 的 nil 指针。

#### 删除行

要删除指定行索引处的行，请调用 `Sheet` 的 `RemoveRowAtIndex(index int)`。此函数仅返回错误代码。

#### 遍历行

`xlsx.Sheet` 提供了一个回调函数来遍历工作表中的每一行，`ForEachRow()`。参数是一个"_行访问函数_"；一个接收指向行的指针作为其唯一参数并返回错误代码的函数。当然，您可以自由使用匿名函数，但为了清晰起见，我在下面的示例中定义了一个名为 `rowVisitor()` 的函数：

```go
func rowVisitor(r *xlsx.Row) error {
    fmt.Println(r)
    return nil
}

func rowStuff() {
    filename := "samplefile.xlsx"
    wb, err := xlsx.OpenFile(filename)
    if err != nil {
        panic(err)
    }
    sh, ok := wb.Sheet["Sample"]
    if !ok {
        panic(errors.New("Sheet not found"))
    }
    fmt.Println("Max row is", sh.MaxRow)
    err = sh.ForEachRow(rowVisitor)
    fmt.Println("Err=", err)
}
```

输出应该类似于下面的控制台日志：

```
== xlsx package tutorial ==
Max row is 4
&{false 0xc00022eb40 0 0 false 0 4 [0xc000294cc0 0xc00022ec00 0xc00022ecc0 0xc00022ed80]}
&{false 0xc00022eb40 0 0 false 1 4 [0xc00022ee40 0xc00022ef00 0xc00022efc0 0xc00022f080]}
&{false 0xc00022eb40 0 0 false 2 4 [0xc00022f140 0xc00022f200 0xc00022f2c0 0xc00022f380]}
&{false 0xc00022eb40 0 0 false 3 4 [0xc00022f440 0xc00022f500 0xc00022f5c0 0xc00022f680]}
Err= <nil>
```

**注意：** 如果您使用的版本*早于* `v3.2.0`，在使用 `ForEachRow()` 时无法知道您当前接收的是*哪一行*（就行号而言）。从 `v.3.2.0` 开始，您可以使用 `Row` 结构体的 `GetCoordinate()` 函数，它将返回一个包含从零开始的行索引的整数。

#### 向行添加单元格

要向现有行追加新单元格，请使用 `AddCell()` 函数。这将返回指向新 `Cell` 的指针（我找不到错误检查，看您是否已达到 xlsx 文件的最大单元格数）。

### 单元格

> 如果您只知道 Excel，那么每个问题看起来都像行和列。 +
> -- _我在需求研讨会上的发言_

单元格是任何电子表格的核心。`xlsx` 包提供了访问、创建和更改单元格的方法，这些将在本章中讨论。在我们开始之前，让我介绍一些在处理电子表格时经常需要的便捷辅助函数。

**提示：** 在 Excel 中引用单元格或单元格范围有两种方法：使用 `A1` 表示法或使用 `RnCn` 表示法。我将在本教程中使用 `A1` 表示法，但如果您有一个小时的时间并想了解为什么 `RnCn` 表示法是 Excel 的魔力所在，请前往 YouTube 观看 Joel Spolsky（前 Excel 程序经理，《Joel on Software》的作者，Trello 的创建者和 Stack Overflow 的联合创始人 – 这足以让您好奇 😉）的这个视频：https://www.youtube.com/watch?v=0nbkaYsR94c[视频 "You suck at Excel"]

如何将像 `A` 或 `BY` 这样的列字母转换为从零开始的列索引？或者如何将像 `BY13` 这样的单元格地址转换为笛卡尔坐标？幸运的是，包包含一些辅助函数。

* `ColIndexToLetters(index int)` – 将数字索引转换为单元格地址的字母组合。
* `ColLettersToIndex(colLetter string)` – 将列地址转换为数字索引。
* `GetCoordsFromCellIDString(cellAddr string)` – 将单元格地址字符串转换为行/列坐标。
* `GetCellIDStringFromCoords(x, y int)` – 将坐标值转换为单元格地址

可以从 `Sheet` 结构体以及 `Row` 结构体访问单个单元格。

#### 从行获取单元格

`GetCell(colIdx int)` 函数返回给定列索引处的 Cell 指针，如果不存在则创建它。这就是没有错误代码的原因。如果您尝试访问"太靠右"的单元格，包将简单地扩展行并为您创建单元格。

如果您想手动添加单元格，可以通过调用 `xlsx.Row` 的 `AddCell()` 函数来实现。这将返回指向新创建的 `xlsx.Cell` 结构体的指针，该结构体已追加到您调用函数的行。

#### 从工作表获取单元格

要从 `Sheet` 结构体获取指向单元格的指针（和错误代码），请使用 `Cell(row, col int)` 函数。在内部，这将调用 Row 的 `GetCell()` 函数，并且它还会扩展工作表以匹配您的坐标。因此，如果您需要知道工作表的数据范围，请确保检查 `MaxRow` 以及 `MaxCol`。

#### 遍历单元格

`Row` 提供了一个回调函数来遍历工作表中的每一行，`ForEachCell()`。参数是一个"_单元格访问函数_"。这是一个接收指向单元格的指针作为其唯一参数并返回错误代码的函数。当然，您可以自由使用匿名函数，但为了清晰起见，我在下面的示例中定义了一个名为 `cellVisitor()` 的函数。这是从示例文件转储工作表（非常简化）的完整列表：

```go
package main

import (
    "errors"
    "fmt"

    "github.com/Denghaoran7617/xlsx"
)

func cellVisitor(c *xlsx.Cell) error {
    value, err := c.FormattedValue()
    if err != nil {
        fmt.Println(err.Error())
    } else {
        fmt.Println("Cell value:", value)
    }
    return err
}

func rowVisitor(r *xlsx.Row) error {
    return r.ForEachCell(cellVisitor)
}

func rowStuff() {
    filename := "samplefile.xlsx"
    wb, err := xlsx.OpenFile(filename)
    if err != nil {
        panic(err)
    }
    sh, ok := wb.Sheet["Sample"]
    if !ok {
        panic(errors.New("Sheet not found"))
    }
    fmt.Println("Max row is", sh.MaxRow)
    sh.ForEachRow(rowVisitor)
}

func main() {
    fmt.Println("== xlsx package tutorial ==")
    rowStuff()
}
```

如果您没有更改示例文件，输出应该如下所示：

```
== xlsx package tutorial ==
Max row is 4
Cell value: Name
Cell value: Date
Cell value: Number
Cell value: Formula
Cell value: Alice
Cell value: 03-20-20
Cell value: 24315
Cell value:  1215.75 €
Cell value: Bob
Cell value: 04-17-20
Cell value: 21345
Cell value:  1067.25 €
Cell value: Charlie
Cell value: 02-08-20
Cell value: 32145
Cell value:  1607.25 €
```

**注意：** 如果您使用的版本*早于* `v3.2.0`，在使用 `ForEachCell()` 时无法知道您当前接收的是*哪个单元格*（就列号和行号而言）。从 `v.3.2.0` 开始，您可以使用 `Cell` 结构体的 `GetCoordinates()` 函数，它将返回一个包含从零开始的列索引和行索引的整数对。

### 单元格类型和内容

#### 单元格类型

Excel 单元格的基本数据类型是

* Bool（布尔值）
* String（字符串）
* Formula（公式）
* Number（数字）
* Date（日期）
* Error（错误）
* Empty（空）

`xlsx.Cell` 为各种数据类型提供了 `SetXXX()` 函数（还将数字数据拆分为 `SetInt()`、`SetFloat()` 等）。

日期值存储为应用了日期格式的数值。是的，上面的列表包含 `Date` 类型，但让我引用代码中的注释：

```go
// d (Date): Cell contains a date in the ISO 8601 format.
// That is the only mention of this format in the XLSX spec.
// Date seems to be unused by the current version of Excel,
// it stores dates as Numeric cells with a date format string.
// For now these cells will have their value output directly.
// It is unclear if the value is supposed to be parsed
// into a number and then formatted using the formatting or not.
```

### 获取单元格值

您可以使用以下函数检索单元格的内容

* `Value()` – 返回字符串
* `FormattedValue()` – 返回应用了单元格格式的值和错误代码
* `String()` – 返回单元格的值作为字符串
* `Formula()` – 返回包含单元格公式的字符串（如果没有公式，则返回空字符串）
* `Int()` - 返回单元格的内容作为整数和错误代码
* `Float()` - 返回单元格的内容作为 float64 和错误代码
* `Bool()` - 返回 `true` 或 `false`
  * 如果单元格具有 `CellTypeBool` 且值等于 `1`，返回 `true`
  * 如果单元格具有 `CellTypeNumeric` 且值非零，返回 `true`
  * 否则，如果 `Value()` 的结果是非空字符串，返回 `true`

### 单元格里有什么？

通常您需要找出单元格的内容，因为仅单元格类型是不够的。为什么不够？让我们看看。示例文件包含一个名为 "Sample" 的工作表，内容如下所示。

|     | A       | B         | C      | D |
|-----|---------|-----------|--------|---|
| **1** | Name    | Date      | Number | Formula |
| **2** | Alice   | 3/20/2020 | 24315  | 1215.75 € |
| **3** | Bob     | 4/17/2020 | 21345  | 1067.25 € |
| **4** | Charlie | 2/8/2020  | 32145  | 1607.25 € |

我们将查看单元格 `D2`（即第 1 行，第 3 列）。下面的示例代码读取单元格并输出使用上一章函数检索的单元格内容。

```go
// 让 sh 是对 xlsx.Sheet 的引用

// 获取 D1 中的单元格，即第 0 行，第 3 列
theCell, err := sh.Cell(0, 3)
if err != nil {
    panic(err)
}
// 我们得到了一个单元格，但里面有什么？
fv, err := theCell.FormattedValue()
if err != nil {
    panic(err)
}
fmt.Println("Numeric cell?:", theCell.Type() == xlsx.CellTypeNumeric)
fmt.Println("String:", theCell.String())
fmt.Println("Formatted:", fv)
fmt.Println("Formula:", theCell.Formula())
```

您应该得到如下输出：

```
Numeric cell?: true
String:  1215.75 €
Formatted:  1215.75 €
Formula: C2*0.05
```

如您所见，为单元格调用 `Type()` 返回"_我是数字类型_"。这很好，但不是全部真相，因为单元格实际上包含一个公式。公式显示在输出的最后一行。如果您有一个"_真正的_"仅包含数字的数字单元格，调用 `Formula()` 的结果是空字符串。因此，如果您想区分这些，请检查单元格的公式是否为空。然后数字单元格才是真正的数字单元格。

### 设置单元格值

要将单元格设置为指定值，请使用 `Cell` 的 `SetXXX()` 函数之一。例如，如果您想为单元格输入公式，请使用 `SetFormula()` 函数并提供公式作为字符串参数。

### 格式化单元格

在格式化方面，我们必须区分显示格式和样式信息（如字体、颜色等）。单元格内容的对齐也属于样式信息。

#### 数字和日期格式

要检索数字（或日期）单元格的格式字符串，请使用 `GetNumberFormat()` 函数，该函数将返回包含当前格式信息的字符串。可以使用 `SetFormat()` 函数（*这里函数名中没有 "Number"*）并提供包含格式信息的字符串来设置格式。

为了简化操作，有一些函数可以同时设置值和格式，例如 `SetFloatWithFormat(val float64, fmt string)`，这样您就不必进行两次函数调用。甚至有一个名为 `NumFmt` 的导出字段可以直接分配格式（`SetFormat()` 基本上只是设置 `NumFmt` 字段）。

Excel 有一整套可以引用的内置格式。有关已知值的列表，请查看 `tealeg/xlsx` 包的仓库，网址为：https://github.com/tealeg/xlsx/blob/master/xmlStyle.go。当然，您也可以使用相同的格式字符串，并使用 `...WithFormat()` 函数之一或 `SetFormat()` 直接设置格式。

让我们为包含在 `c` 中的单元格设置一个数字格式，该格式将以红色显示负值并使用两位小数精度：

```go
c.NumFmt = "#0.00;[RED]-#0.00"

// 或者您可以使用
c.SetFormat("#0.00;[RED]-#0.00")
```

**注意：** `xlsx.File` 结构体有一个导出的字段 `Date1904`。在大多数 xlsx 文件中，该值应为 `false`，这意味着"_基准日期_"是 1900 年 1 月 1 日。如前所述，Excel 将日期存储为数值（自"_基准日期_"以来经过的天数）。如果 `Date1904` 的值为 `true`，则"_基准日期_"是 1904 年 1 月 1 日。这样做的原因是早期 Macintosh 版 Excel 的日期处理问题，因为 1900 年*不是*闰年。这里的 `tealeg/xlsx` 包会自动处理此问题，因此应该不需要担心。但如果您确实使用自己的例程处理日期，您应该检查哪个日期是"第零天"。您可以在 Microsoft 网站的 https://docs.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system[此 Excel 支持文档] 中找到有关此主题的更详细信息。

### 样式

样式提供有关单元格布局和装饰各个方面的信息，可以由多个单元格使用。虽然您*可以*为每个单元格应用新样式，但这并不意味着您*应该*这样做。为什么要使用 300 个包含相同信息的对象？最好创建一个样式并重复使用它。样式中有什么？

```go
// Style 是一个高级结构体，旨在为用户提供
// 对 XLSX 文件中样式内容的访问。
type Style struct {
    Border          Border
    Fill            Fill
    Font            Font
    ApplyBorder     bool
    ApplyFill       bool
    ApplyFont       bool
    ApplyAlignment  bool
    Alignment       Alignment
    NamedStyleIndex *int
}
```

#### 分配样式

让我们创建一个样式！

```go
myStyle := xlsx.NewStyle()
```

很简单，不是吗？好的，这返回一个指向空样式的指针，因此我们必须将某些字段设置为有用的值：

```go
myStyle := xlsx.NewStyle()
myStyle.Alignment.Horizontal = "right"
myStyle.Fill.FgColor = "FFFFFF00"
myStyle.Fill.PatternType = "solid"
myStyle.Font.Name = "Georgia"
myStyle.Font.Size = 11
myStyle.Font.Bold = true
myStyle.ApplyAlignment = true
myStyle.ApplyFill = true
myStyle.ApplyFont = true
```

现在我们有了样式，我们可以使用以下语句将此样式分配给单元格（我将使用 `aCell` 作为单元格变量）：`aCell.SetStyle(myStyle)`。稍后，您将看到列也有 `SetStyle()` 函数。

#### 检索样式信息

使用单元格的 `GetStyle()` 函数返回指向 `Style` 结构体的指针。如果您从未更改样式，返回的样式将是工作表的默认样式。下面的代码读取文件 `samplefile.xlsx` 中名为 _Styles_ 的工作表的单元格 0, 1（_这是 A2_）并显示一些可用的样式信息。_请注意，为了简洁起见，没有错误检查。这在演示代码中是可以的，但不要在生产环境中这样做。_ 😉

```go
package main

import (
    "errors"
    "fmt"

    "github.com/Denghaoran7617/xlsx"
    )

func MAIN() {
    filename := "samplefile.xlsx"
    wb, _ := xlsx.OpenFile(filename)
    sh := wb.Sheet["Styles"]
    cell, _ := sh.Cell(0, 1)
    style := cell.GetStyle()
    fmt.Println("Cell value:", cell.String())
    fmt.Println("Font:", style.Font.Name)
    fmt.Println("Size:", style.Font.Size)
    fmt.Println("H-Align:", style.Alignment.Horizontal)
    fmt.Println("ForeColor:", style.Fill.FgColor)
    fmt.Println("BackColor:", style.Fill.BgColor)
}
```

## 使用列

如果 `xlsx` 包中有一个主题在主要版本中发生了**很大**变化，那就是列。因此，让我们看看从 V3 开始的工作方式。我个人强烈建议仅为了列功能就升级到包的 V3，因为它现在更接近 Excel 文件的内部工作原理。

### 定义列

列结构体*不*与工作表中的单列单元格相关。相反，工作表至少有一个列定义，可以与每一列关联。工作表的列定义的最大数量当然等于工作表中的列数（然后我们将有列定义和工作表列的 1:1 关联）。

这是来自仓库的 `Col` 结构体的定义：

```go
type Col struct {
    Min          int
    Max          int
    Hidden       *bool
    Width        *float64
    Collapsed    *bool
    OutlineLevel *uint8
    BestFit      *bool
    CustomWidth  *bool
    Phonetic     *bool
    // contains filtered or unexported fields
}
```

您将看到有两个字段 `Min` 和 `Max`，它们定义此 `Col` 将关联的工作表列的范围。有一个名为 `NewColForRange()` 的函数，它接受两个参数（min 和 max）并返回指向 `Col` 结构体的指针。除非您设置某些字段并使用 `Sheet` 结构体的 `SetColParameters()` 函数将此列与工作表关联，否则这还不是很有用。

下面的代码片段创建 `Col` 定义，设置宽度并分配样式。然后我们调用工作表的 `SetColParameters()` 函数将此列与工作表关联。列 A 到 E 中的任何单元格将具有 12.5 的宽度并使用 `myStyle` 指针引用的样式（见上文）。

```go
// 创建与工作表列 A 到 E（索引 1 到 5）相关的列
newColumn := NewColForRange(1,5)
newColumn.SetWidth(12.5)
// 我们在上面定义了一个样式，所以让我们将此样式分配给列的所有单元格
newColumn.SetStyle(myStyle)
// 现在将工作表与此列关联
sh.SetColParameters(newColumn)
```

如您所见，我们可以写入工作表列 A 到 E 中的任何单元格，但只有一个列定义。当然，您可以创建五列，每列对应一个工作表列。例如，如果您需要五种不同的样式或五个不同的宽度值，这就是要走的路。顺便说一句，如果您创建新的 `Col` 结构体并在工作表中使用它们，包会处理插入、删除或为新列让路。

### 宽度单位

让我们想象您将列的宽度设置为值 '12.5'。这意味着什么？既不是英寸也不是像素。xlsx 文件中的列宽表示为以正常样式字体呈现的数字 0-9 的最大数字宽度的字符数。值 `12.5` 意味着（假设字体中 0-9 的每个数字具有相同的宽度）12.5 个数字将适合该列的单元格。脚注：[即使在比例字体中，大多数时候数字使用相同的宽度，以使表格中的数值更易于阅读。]

### 设置列的宽度

您可以直接使用工作表的 `SetColWidth` 函数设置列范围的宽度。此函数具有以下签名：

```go
func (s *Sheet) SetColWidth(min, max int, width float64)
```

如果您需要设置单列的宽度，请为 `min` 和 `max` 指定相同的值。

在处理列结构体时，您可以使用 `Column` 结构体的 `SetWidth` 函数来设置链接到此列的所有单元格的宽度。该函数接受一个参数，宽度作为 `float64`。

## 其他工作簿内容

### 将内容获取为字节

也许您想以特殊方式处理工作的结果，而不是将 `.xlsx` 文件写入磁盘。`xlsx.File` 结构体有一个附加的 `Write()` 方法，它将文件写入任何 `io.Writer`。请参阅下面的示例，了解如何将 xlsx 文件作为字节缓冲区获取。

```go
file := xlsx.NewFile()
/*
    对文件做一些操作...
*/
var b bytes.Buffer
writer := bufio.NewWriter(&b)
file.Write(writer)
theBytes := b.Bytes()
/*
    现在您在 b 中有字节流。
    如果您使用满足 Writer 接口的其他类型，请继续。
*/
```

### 定义的名称（命名范围）

您可以为单元格或单元格范围定义名称。此名称可用于公式中，使内容更易于阅读和理解。此信息存储在 Excel 文件的 `definedName` 元素中。您可以使用 `xlsx.File` 结构体的 `DefinedNames` 字段访问此定义的名称列表。它保存指向 `DefinedName` 结构体的指针切片（`[]*xlsx.xlsxDefinedName`）。有几个字段，您可以在 https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.definedname.aspx[此 MSDN 文档] 中获取更详细的信息。对于我们的目的，使用 `Name` 和 `Data` 就足够了。

* `Name` 是单元格或单元格范围的名称字符串。通常，名称解释此名称引用的对象的用途，使查找和使用此对象更容易。
* `Data` 包含对单元格或单元格范围的引用的字符串

定义名称受一些语法规则约束。向 https://docs.devexpress.com/WindowsForms/14691/Controls-and-Libraries/Spreadsheet/Defined-Names#syntax-rules-for-names[DevExpress] 致意，感谢这些信息！

* 名称必须以字母或下划线开头，最小长度为 1 个字符。
* 名称的剩余字符可以是字母、下划线、数字或句点。
* 不能将单个字母 `C`、`c`、`R` 或 `r` 用作定义的名称。
* 名称不能与单元格引用相同（例如，`A1`、`$M$15`）。
* 名称不能包含空格（请使用下划线符号和句点代替）。
* 名称的长度不能超过 255 个字符。
* 名称不区分大小写。

下面列出了一些 `Data` 的示例：

* `Sample!$A$2` – 引用名为 "Sample" 的工作表中的单个单元格 A2
* `Styles!$A$2:$A$8` – 引用名为 "Styles" 的工作表中从 A2 到 A8 的范围
* `Sheet1!$D$20` – 引用名为 "Sheet1" 的工作表上的单元格 D20

