# XLSX - Go 语言 Excel 文件处理库

## 项目概述

XLSX 是一个用于简化读取和写入 Microsoft Excel XLSX 格式文件的 Go 语言库。XLSX 格式是自 2002 年以来 Microsoft Excel 使用的 XML 格式。该库提供了完整的 Excel 文件读写功能，支持单元格、行、列、工作表等基本操作，以及样式、格式化、数据验证、超链接等高级特性。

### 主要特性

- **文件操作**：支持创建、打开、读取和保存 XLSX 文件
- **数据结构**：提供 File、Sheet、Row、Cell 等高级数据结构
- **样式支持**：支持字体、边框、填充、对齐等样式设置
- **数据类型**：支持字符串、数字、布尔值、日期时间等基本类型
- **复杂类型**：支持 map、slice、array 等复杂类型的读取（通过 JSON 格式）
- **单元格存储**：支持内存存储和磁盘存储两种方式，适用于大文件处理
- **数据验证**：支持单元格数据验证规则
- **超链接**：支持单元格超链接功能
- **合并单元格**：支持单元格合并操作

## 安装

```bash
go get github.com/Denghaoran7617/xlsx
```

## 快速开始

### 创建 Excel 文件

```go
package main

import (
    "github.com/Denghaoran7617/xlsx"
)

func main() {
    file := xlsx.NewFile()
    sheet, _ := file.AddSheet("Sheet1")
    row := sheet.AddRow()
    cell := row.AddCell()
    cell.Value = "Hello"
    cell = row.AddCell()
    cell.Value = "World"
    file.Save("example.xlsx")
}
```

### 读取 Excel 文件

```go
package main

import (
    "fmt"
    "github.com/Denghaoran7617/xlsx"
)

func main() {
    file, err := xlsx.OpenFile("example.xlsx")
    if err != nil {
        panic(err)
    }
    
    for _, sheet := range file.Sheets {
        fmt.Printf("Sheet: %s\n", sheet.Name)
        sheet.ForEachRow(func(row *xlsx.Row) error {
            row.ForEachCell(func(cell *xlsx.Cell) error {
                value, _ := cell.FormattedValue()
                fmt.Printf("%s\t", value)
                return nil
            })
            fmt.Println()
            return nil
        })
    }
}
```

## API 文档

### 文件操作

#### NewFile
创建新的 File 结构体。

**函数签名：**
```go
func NewFile(options ...FileOption) *File
```

**参数：**
- `options` (可选): FileOption 函数，用于配置文件行为

**返回：**
- `*File`: 新创建的 File 指针

#### OpenFile
打开指定的 XLSX 文件并返回 File 结构体。

**函数签名：**
```go
func OpenFile(fileName string, options ...FileOption) (file *File, err error)
```

**参数：**
- `fileName` (string): XLSX 文件路径
- `options` (可选): FileOption 函数，用于配置文件行为

**返回：**
- `file` (*File): 打开的 File 指针
- `err` (error): 错误信息

#### OpenBinary
从字节数组打开 XLSX 文件。

**函数签名：**
```go
func OpenBinary(bs []byte, options ...FileOption) (*File, error)
```

**参数：**
- `bs` ([]byte): XLSX 文件的字节数组
- `options` (可选): FileOption 函数

**返回：**
- `*File`: File 指针
- `error`: 错误信息

#### OpenReaderAt
从 io.ReaderAt 打开 XLSX 文件。

**函数签名：**
```go
func OpenReaderAt(r io.ReaderAt, size int64, options ...FileOption) (*File, error)
```

**参数：**
- `r` (io.ReaderAt): 读取器
- `size` (int64): 文件大小
- `options` (可选): FileOption 函数

**返回：**
- `*File`: File 指针
- `error`: 错误信息

#### File.Save
将 File 保存到指定路径的 xlsx 文件。

**方法签名：**
```go
func (f *File) Save(path string) (err error)
```

**参数：**
- `path` (string): 保存路径

**返回：**
- `err` (error): 错误信息

#### File.Write
将 File 写入 io.Writer 作为 xlsx 格式。

**方法签名：**
```go
func (f *File) Write(writer io.Writer) error
```

**参数：**
- `writer` (io.Writer): 写入目标

**返回：**
- `error`: 错误信息

#### File.AddSheet
向 File 添加新的 Sheet。

**方法签名：**
```go
func (f *File) AddSheet(sheetName string) (*Sheet, error)
```

**参数：**
- `sheetName` (string): Sheet 名称（1-31 个字符，不能包含 : \ / ? * [ ]）

**返回：**
- `*Sheet`: Sheet 指针
- `error`: 错误信息

#### FileToSlice
返回 Excel XLSX 文件中的原始数据作为三维切片。

**函数签名：**
```go
func FileToSlice(path string, options ...FileOption) ([][][]string, error)
```

**参数：**
- `path` (string): 文件路径
- `options` (可选): FileOption 函数

**返回：**
- `[][][]string`: 三维字符串切片（sheet -> row -> cell）
- `error`: 错误信息

#### FileToSliceUnmerged
返回 Excel XLSX 文件中的原始数据，合并单元格将被拆分。

**函数签名：**
```go
func FileToSliceUnmerged(path string, options ...FileOption) ([][][]string, error)
```

**参数：**
- `path` (string): 文件路径
- `options` (可选): FileOption 函数

**返回：**
- `[][][]string`: 三维字符串切片
- `error`: 错误信息

### 文件选项

#### RowLimit
限制任何给定 Sheet 中处理的行数为前 n 行。

**函数签名：**
```go
func RowLimit(n int) FileOption
```

**参数：**
- `n` (int): 行数限制

**返回：**
- `FileOption`: 文件选项函数

#### ColLimit
限制任何给定 Sheet 中处理的列数为前 n 列。

**函数签名：**
```go
func ColLimit(n int) FileOption
```

**参数：**
- `n` (int): 列数限制

**返回：**
- `FileOption`: 文件选项函数

#### ValueOnly
将所有 NULL 值视为无意义，在解码 worksheet.xml 之前删除所有 NULL 值单元格。

**函数签名：**
```go
func ValueOnly() FileOption
```

**返回：**
- `FileOption`: 文件选项函数

### Sheet 操作

#### NewSheet
使用默认 CellStore 构造 Sheet 并返回指针。

**函数签名：**
```go
func NewSheet(name string) (*Sheet, error)
```

**参数：**
- `name` (string): Sheet 名称

**返回：**
- `*Sheet`: Sheet 指针
- `error`: 错误信息

#### NewSheetWithCellStore
使用指定的 CellStore 构造函数构造 Sheet。

**函数签名：**
```go
func NewSheetWithCellStore(name string, constructor CellStoreConstructor) (*Sheet, error)
```

**参数：**
- `name` (string): Sheet 名称
- `constructor` (CellStoreConstructor): CellStore 构造函数

**返回：**
- `*Sheet`: Sheet 指针
- `error`: 错误信息

#### Sheet.Close
移除 Sheet 的依赖资源。

**方法签名：**
```go
func (s *Sheet) Close()
```

#### Sheet.AddRow
向 Sheet 添加新行。

**方法签名：**
```go
func (s *Sheet) AddRow() *Row
```

**返回：**
- `*Row`: 新创建的 Row 指针

#### Sheet.AddRowAtIndex
在指定索引处向 Sheet 添加新行。

**方法签名：**
```go
func (s *Sheet) AddRowAtIndex(index int) (*Row, error)
```

**参数：**
- `index` (int): 行索引（从 0 开始）

**返回：**
- `*Row`: 新创建的 Row 指针
- `error`: 错误信息

#### Sheet.RemoveRowAtIndex
移除指定索引处的行。

**方法签名：**
```go
func (s *Sheet) RemoveRowAtIndex(index int) error
```

**参数：**
- `index` (int): 行索引

**返回：**
- `error`: 错误信息

#### Sheet.Row
获取指定索引的行。

**方法签名：**
```go
func (s *Sheet) Row(idx int) (*Row, error)
```

**参数：**
- `idx` (int): 行索引（从 0 开始）

**返回：**
- `*Row`: Row 指针
- `error`: 错误信息

#### Sheet.Cell
通过行列坐标获取单元格。

**方法签名：**
```go
func (s *Sheet) Cell(row, col int) (*Cell, error)
```

**参数：**
- `row` (int): 行索引（从 0 开始）
- `col` (int): 列索引（从 0 开始）

**返回：**
- `*Cell`: Cell 指针
- `error`: 错误信息

#### Sheet.Col
返回应用于此列索引的 Col，如果不存在则返回 nil。

**方法签名：**
```go
func (s *Sheet) Col(idx int) *Col
```

**参数：**
- `idx` (int): 列索引（从 0 开始，列号从 1 开始）

**返回：**
- `*Col`: Col 指针

#### Sheet.ForEachRow
遍历 Sheet 中的每一行，调用提供的 RowVisitor 函数。

**方法签名：**
```go
func (s *Sheet) ForEachRow(rv RowVisitor, options ...RowVisitorOption) error
```

**参数：**
- `rv` (RowVisitor): 行访问函数
- `options` (可选): RowVisitorOption 选项

**返回：**
- `error`: 错误信息

#### Sheet.SetColWidth
设置列范围宽度。

**方法签名：**
```go
func (s *Sheet) SetColWidth(min, max int, width float64)
```

**参数：**
- `min` (int): 最小列号（从 1 开始）
- `max` (int): 最大列号（从 1 开始）
- `width` (float64): 列宽

#### Sheet.SetColAutoWidth
根据最大单元格内容尝试猜测列的最佳宽度。

**方法签名：**
```go
func (s *Sheet) SetColAutoWidth(colIndex int, width func(string) float64) error
```

**参数：**
- `colIndex` (int): 列索引（从 1 开始）
- `width` (func(string) float64): 宽度计算函数

**返回：**
- `error`: 错误信息

#### Sheet.SetColParameters
设置列的参数。

**方法签名：**
```go
func (s *Sheet) SetColParameters(col *Col)
```

**参数：**
- `col` (*Col): Col 结构体指针

#### Sheet.SetOutlineLevel
设置列范围的大纲级别。

**方法签名：**
```go
func (s *Sheet) SetOutlineLevel(minCol, maxCol int, outlineLevel uint8)
```

**参数：**
- `minCol` (int): 最小列号
- `maxCol` (int): 最大列号
- `outlineLevel` (uint8): 大纲级别

#### Sheet.SetType
设置列范围的类型。

**方法签名：**
```go
func (s *Sheet) SetType(minCol, maxCol int, cellType CellType)
```

**参数：**
- `minCol` (int): 最小列号
- `maxCol` (int): 最大列号
- `cellType` (CellType): 单元格类型

#### Sheet.AddDataValidation
向单元格范围添加数据验证。

**方法签名：**
```go
func (s *Sheet) AddDataValidation(dv *xlsxDataValidation)
```

**参数：**
- `dv` (*xlsxDataValidation): 数据验证结构体

#### IsSaneSheetName
检查 Sheet 名称是否有效。

**函数签名：**
```go
func IsSaneSheetName(sheetName string) error
```

**参数：**
- `sheetName` (string): Sheet 名称

**返回：**
- `error`: 错误信息（如果名称无效）

### Row 操作

#### Row.GetCoordinate
返回行的 y 坐标（行号，从 0 开始）。

**方法签名：**
```go
func (r *Row) GetCoordinate() int
```

**返回：**
- `int`: 行号（从 0 开始）

#### Row.SetHeight
设置行高（PostScript 点）。

**方法签名：**
```go
func (r *Row) SetHeight(ht float64)
```

**参数：**
- `ht` (float64): 高度（PostScript 点）

#### Row.SetHeightCM
设置行高（厘米）。

**方法签名：**
```go
func (r *Row) SetHeightCM(ht float64)
```

**参数：**
- `ht` (float64): 高度（厘米）

#### Row.GetHeight
返回行高（PostScript 点）。

**方法签名：**
```go
func (r *Row) GetHeight() float64
```

**返回：**
- `float64`: 高度（PostScript 点）

#### Row.SetOutlineLevel
设置行的大纲级别（用于折叠行）。

**方法签名：**
```go
func (r *Row) SetOutlineLevel(outlineLevel uint8)
```

**参数：**
- `outlineLevel` (uint8): 大纲级别

#### Row.GetOutlineLevel
返回行的大纲级别。

**方法签名：**
```go
func (r *Row) GetOutlineLevel() uint8
```

**返回：**
- `uint8`: 大纲级别

#### Row.AddCell
向行末尾添加新单元格。

**方法签名：**
```go
func (r *Row) AddCell() *Cell
```

**返回：**
- `*Cell`: 新创建的 Cell 指针

#### Row.PushCell
向行末尾添加预定义的单元格。

**方法签名：**
```go
func (r *Row) PushCell(c *Cell)
```

**参数：**
- `c` (*Cell): Cell 指针

#### Row.GetCell
返回指定列索引的单元格，如果不存在则创建。

**方法签名：**
```go
func (r *Row) GetCell(colIdx int) *Cell
```

**参数：**
- `colIdx` (int): 列索引（从 0 开始）

**返回：**
- `*Cell`: Cell 指针

#### Row.ForEachCell
遍历行中的每个已定义单元格，调用提供的 CellVisitorFunc。

**方法签名：**
```go
func (r *Row) ForEachCell(cvf CellVisitorFunc, option ...CellVisitorOption) error
```

**参数：**
- `cvf` (CellVisitorFunc): 单元格访问函数
- `option` (可选): CellVisitorOption 选项

**返回：**
- `error`: 错误信息

#### Row.WriteSlice
将切片写入行。

**方法签名：**
```go
func (r *Row) WriteSlice(e interface{}, cols int) int
```

**参数：**
- `e` (interface{}): 切片或切片指针
- `cols` (int): 要写入的列数（< 0 表示全部）

**返回：**
- `int`: 写入的列数，如果不是切片类型返回 -1

#### Row.WriteStruct
将结构体写入行。

**方法签名：**
```go
func (r *Row) WriteStruct(e interface{}, cols int) int
```

**参数：**
- `e` (interface{}): 指向结构体的指针
- `cols` (int): 要写入的列数（< 0 表示全部）

**返回：**
- `int`: 写入的列数，如果不是结构体指针返回 -1

#### Row.ReadStruct
从行读取结构体到指针。

**方法签名：**
```go
func (r *Row) ReadStruct(ptr interface{}) error
```

**参数：**
- `ptr` (interface{}): 指向结构体的指针，结构体字段需要 `xlsx:"N"` 标签，其中 N 是单元格索引

**返回：**
- `error`: 错误信息

**支持的类型：**
- 基本类型：string、int、int8、int16、int32、int64、float64、bool
- 时间类型：time.Time
- 复杂类型：map、slice、array（通过 parseValue 转换，支持 JSON 格式）

### Cell 操作

#### Cell.String
返回单元格的字符串表示。

**方法签名：**
```go
func (c *Cell) String() string
```

**返回：**
- `string`: 单元格字符串值

#### Cell.FormattedValue
返回格式化后的单元格值。

**方法签名：**
```go
func (c *Cell) FormattedValue() (string, error)
```

**返回：**
- `string`: 格式化后的值
- `error`: 错误信息

#### Cell.Int
返回单元格的整数值。

**方法签名：**
```go
func (c *Cell) Int() (int, error)
```

**返回：**
- `int`: 整数值
- `error`: 错误信息

#### Cell.Int64
返回单元格的 int64 值。

**方法签名：**
```go
func (c *Cell) Int64() (int64, error)
```

**返回：**
- `int64`: int64 值
- `error`: 错误信息

#### Cell.Float
返回单元格的浮点数值。

**方法签名：**
```go
func (c *Cell) Float() (float64, error)
```

**返回：**
- `float64`: 浮点数值
- `error`: 错误信息

#### Cell.Bool
返回单元格的布尔值。

**方法签名：**
```go
func (c *Cell) Bool() bool
```

**返回：**
- `bool`: 布尔值

#### Cell.GetTime
返回单元格的时间值。

**方法签名：**
```go
func (c *Cell) GetTime(allowZeroTime bool) (time.Time, error)
```

**参数：**
- `allowZeroTime` (bool): 是否允许零时间

**返回：**
- `time.Time`: 时间值
- `error`: 错误信息

#### Cell.SetString
设置单元格的字符串值。

**方法签名：**
```go
func (c *Cell) SetString(s string)
```

**参数：**
- `s` (string): 字符串值

#### Cell.SetInt
设置单元格的整数值。

**方法签名：**
```go
func (c *Cell) SetInt(n int)
```

**参数：**
- `n` (int): 整数值

#### Cell.SetInt64
设置单元格的 int64 值。

**方法签名：**
```go
func (c *Cell) SetInt64(n int64)
```

**参数：**
- `n` (int64): int64 值

#### Cell.SetNumeric
设置单元格的数值。

**方法签名：**
```go
func (c *Cell) SetNumeric(s string)
```

**参数：**
- `s` (string): 数值字符串

#### Cell.SetFloat
设置单元格的浮点数值。

**方法签名：**
```go
func (c *Cell) SetFloat(n float64)
```

**参数：**
- `n` (float64): 浮点数值

#### Cell.SetBool
设置单元格的布尔值。

**方法签名：**
```go
func (c *Cell) SetBool(b bool)
```

**参数：**
- `b` (bool): 布尔值

#### Cell.SetDateTime
设置单元格的日期时间值。

**方法签名：**
```go
func (c *Cell) SetDateTime(t time.Time)
```

**参数：**
- `t` (time.Time): 时间值

#### Cell.SetDate
设置单元格的日期值。

**方法签名：**
```go
func (c *Cell) SetDate(t time.Time)
```

**参数：**
- `t` (time.Time): 时间值

#### Cell.SetValue
设置单元格的值（自动类型推断）。

**方法签名：**
```go
func (c *Cell) SetValue(n interface{})
```

**参数：**
- `n` (interface{}): 值（支持多种类型）

#### Cell.SetFormula
设置单元格的公式。

**方法签名：**
```go
func (c *Cell) SetFormula(formula string)
```

**参数：**
- `formula` (string): 公式字符串

#### Cell.SetStringFormula
设置单元格的字符串公式。

**方法签名：**
```go
func (c *Cell) SetStringFormula(formula string)
```

**参数：**
- `formula` (string): 公式字符串

#### Cell.Formula
返回单元格的公式字符串。

**方法签名：**
```go
func (c *Cell) Formula() string
```

**返回：**
- `string`: 公式字符串

#### Cell.GetStyle
返回与单元格关联的样式。

**方法签名：**
```go
func (c *Cell) GetStyle() *Style
```

**返回：**
- `*Style`: Style 指针

#### Cell.SetStyle
设置单元格的样式。

**方法签名：**
```go
func (c *Cell) SetStyle(style *Style)
```

**参数：**
- `style` (*Style): Style 指针

#### Cell.GetNumberFormat
返回单元格的数字格式字符串。

**方法签名：**
```go
func (c *Cell) GetNumberFormat() string
```

**返回：**
- `string`: 数字格式字符串

#### Cell.GetCoordinates
返回单元格的坐标对（列号，行号）。

**方法签名：**
```go
func (c *Cell) GetCoordinates() (int, int)
```

**返回：**
- `int`: 列号（从 0 开始）
- `int`: 行号（从 0 开始）

#### Cell.SetDataValidation
设置数据验证。

**方法签名：**
```go
func (c *Cell) SetDataValidation(dd *xlsxDataValidation)
```

**参数：**
- `dd` (*xlsxDataValidation): 数据验证结构体

#### Cell.Modified
返回单元格自上次持久化以来是否已被修改。

**方法签名：**
```go
func (c *Cell) Modified() bool
```

**返回：**
- `bool`: 是否已修改

### Col 操作

#### NewColForRange
返回指向新 Col 的指针，该 Col 将应用于 min 到 max（包含）范围内的列。

**函数签名：**
```go
func NewColForRange(min, max int) *Col
```

**参数：**
- `min` (int): 最小列号（从 1 开始）
- `max` (int): 最大列号（从 1 开始）

**返回：**
- `*Col`: Col 指针

#### Col.SetWidth
设置应用此 Col 的列的宽度。

**方法签名：**
```go
func (c *Col) SetWidth(width float64)
```

**参数：**
- `width` (float64): 列宽

#### Col.SetType
根据要设置的类型设置列的格式字符串。

**方法签名：**
```go
func (c *Col) SetType(cellType CellType)
```

**参数：**
- `cellType` (CellType): 单元格类型

#### Col.GetStyle
返回与 Col 关联的样式。

**方法签名：**
```go
func (c *Col) GetStyle() *Style
```

**返回：**
- `*Style`: Style 指针

#### Col.SetStyle
设置 Col 的样式。

**方法签名：**
```go
func (c *Col) SetStyle(style *Style)
```

**参数：**
- `style` (*Style): Style 指针

#### Col.SetOutlineLevel
设置列的大纲级别。

**方法签名：**
```go
func (c *Col) SetOutlineLevel(outlineLevel uint8)
```

**参数：**
- `outlineLevel` (uint8): 大纲级别

### Style 操作

#### NewStyle
返回使用默认值初始化的新 Style 结构体。

**函数签名：**
```go
func NewStyle() *Style
```

**返回：**
- `*Style`: Style 指针

#### NewBorder
创建新的边框样式。

**函数签名：**
```go
func NewBorder(left, right, top, bottom string) *Border
```

**参数：**
- `left` (string): 左边框样式
- `right` (string): 右边框样式
- `top` (string): 上边框样式
- `bottom` (string): 下边框样式

**返回：**
- `*Border`: Border 指针

#### NewFill
创建新的填充样式。

**函数签名：**
```go
func NewFill(patternType, fgColor, bgColor string) *Fill
```

**参数：**
- `patternType` (string): 填充模式类型
- `fgColor` (string): 前景色
- `bgColor` (string): 背景色

**返回：**
- `*Fill`: Fill 指针

#### NewFont
创建新的字体样式。

**函数签名：**
```go
func NewFont(size float64, name string) *Font
```

**参数：**
- `size` (float64): 字体大小
- `name` (string): 字体名称

**返回：**
- `*Font`: Font 指针

#### SetDefaultFont
设置默认字体。

**函数签名：**
```go
func SetDefaultFont(size float64, name string)
```

**参数：**
- `size` (float64): 字体大小
- `name` (string): 字体名称

#### DefaultFont
返回默认字体。

**函数签名：**
```go
func DefaultFont() *Font
```

**返回：**
- `*Font`: Font 指针

#### DefaultFill
返回默认填充。

**函数签名：**
```go
func DefaultFill() *Fill
```

**返回：**
- `*Fill`: Fill 指针

#### DefaultBorder
返回默认边框。

**函数签名：**
```go
func DefaultBorder() *Border
```

**返回：**
- `*Border`: Border 指针

#### DefaultAlignment
返回默认对齐方式。

**函数签名：**
```go
func DefaultAlignment() *Alignment
```

**返回：**
- `*Alignment`: Alignment 指针

#### DefaultAutoWidth
可用作自动宽度的默认缩放函数。

**函数签名：**
```go
func DefaultAutoWidth(s string) float64
```

**参数：**
- `s` (string): 字符串

**返回：**
- `float64`: 计算出的宽度

### 工具函数

#### ColLettersToIndex
将基于字符的列引用转换为从零开始的数字列标识符。

**函数签名：**
```go
func ColLettersToIndex(letters string) int
```

**参数：**
- `letters` (string): 列字母（如 "A", "B", "AA"）

**返回：**
- `int`: 列索引（从 0 开始）

#### ColIndexToLetters
将基于零的数字列标识符转换为字符代码。

**函数签名：**
```go
func ColIndexToLetters(n int) string
```

**参数：**
- `n` (int): 列索引（从 0 开始）

**返回：**
- `string`: 列字母（如 "A", "B", "AA"）

#### RowIndexToString
将基于零的数字行标识符转换为其字符串表示。

**函数签名：**
```go
func RowIndexToString(rowRef int) string
```

**参数：**
- `rowRef` (int): 行索引（从 0 开始）

**返回：**
- `string`: 行号字符串（从 1 开始）

#### GetCoordsFromCellIDString
从 Excel 格式的单元格名称返回基于零的笛卡尔坐标。

**函数签名：**
```go
func GetCoordsFromCellIDString(cellIDString string) (x, y int, err error)
```

**参数：**
- `cellIDString` (string): 单元格 ID 字符串（如 "A1", "B3"）

**返回：**
- `x` (int): 列索引（从 0 开始）
- `y` (int): 行索引（从 0 开始）
- `err` (error): 错误信息

#### GetCellIDStringFromCoords
返回表示一对基于零的笛卡尔坐标的 Excel 格式单元格名称。

**函数签名：**
```go
func GetCellIDStringFromCoords(x, y int) string
```

**参数：**
- `x` (int): 列索引（从 0 开始）
- `y` (int): 行索引（从 0 开始）

**返回：**
- `string`: 单元格 ID 字符串（如 "A1", "B3"）

#### GetCellIDStringFromCoordsWithFixed
返回表示一对基于零的笛卡尔坐标的 Excel 格式单元格名称，可以指定任一值为固定。

**函数签名：**
```go
func GetCellIDStringFromCoordsWithFixed(x, y int, xFixed, yFixed bool) string
```

**参数：**
- `x` (int): 列索引（从 0 开始）
- `y` (int): 行索引（从 0 开始）
- `xFixed` (bool): 列是否固定
- `yFixed` (bool): 行是否固定

**返回：**
- `string`: 单元格 ID 字符串（如 "$A$1", "B$3"）

### RichText 操作

#### NewRichTextColorFromARGB
从 ARGB 值创建新的富文本颜色。

**函数签名：**
```go
func NewRichTextColorFromARGB(alpha, red, green, blue int) *RichTextColor
```

**参数：**
- `alpha` (int): Alpha 通道值
- `red` (int): 红色通道值
- `green` (int): 绿色通道值
- `blue` (int): 蓝色通道值

**返回：**
- `*RichTextColor`: RichTextColor 指针

#### NewRichTextColorFromThemeColor
从主题颜色创建新的富文本颜色。

**函数签名：**
```go
func NewRichTextColorFromThemeColor(themeColor int) *RichTextColor
```

**参数：**
- `themeColor` (int): 主题颜色索引

**返回：**
- `*RichTextColor`: RichTextColor 指针

### RefTable 操作

#### NewSharedStringRefTable
创建新的共享字符串引用表。

**函数签名：**
```go
func NewSharedStringRefTable(size int) *RefTable
```

**参数：**
- `size` (int): 初始大小

**返回：**
- `*RefTable`: RefTable 指针

#### RefTable.ResolveSharedString
解析共享字符串。

**方法签名：**
```go
func (rt *RefTable) ResolveSharedString(index int) (plainText string, richText []RichTextRun)
```

**参数：**
- `index` (int): 字符串索引

**返回：**
- `plainText` (string): 纯文本
- `richText` ([]RichTextRun): 富文本运行

#### RefTable.AddString
添加字符串到共享字符串表。

**方法签名：**
```go
func (rt *RefTable) AddString(str string) int
```

**参数：**
- `str` (string): 字符串

**返回：**
- `int`: 字符串索引

#### RefTable.AddRichText
添加富文本到共享字符串表。

**方法签名：**
```go
func (rt *RefTable) AddRichText(r []RichTextRun) int
```

**参数：**
- `r` ([]RichTextRun): 富文本运行

**返回：**
- `int`: 字符串索引

#### RefTable.Length
返回共享字符串表的长度。

**方法签名：**
```go
func (rt *RefTable) Length() int
```

**返回：**
- `int`: 长度

### CellStore 操作

#### UseMemoryCellStore
使 File 的所有 Sheet 实例使用 Memory 作为其支持存储。

**函数签名：**
```go
func UseMemoryCellStore(f *File)
```

**参数：**
- `f` (*File): File 指针

#### NewMemoryCellStore
返回基于 Memory 的 CellStore。

**函数签名：**
```go
func NewMemoryCellStore() (CellStore, error)
```

**返回：**
- `CellStore`: CellStore 接口
- `error`: 错误信息

#### UseDiskVCellStore
使 File 的所有 Sheet 实例使用 DiskV 作为其支持存储。

**函数签名：**
```go
func UseDiskVCellStore(f *File)
```

**参数：**
- `f` (*File): File 指针

#### NewDiskVCellStore
返回基于 DiskV 的 CellStore。

**函数签名：**
```go
func NewDiskVCellStore() (CellStore, error)
```

**返回：**
- `CellStore`: CellStore 接口
- `error`: 错误信息

### 访问选项

#### SkipEmptyRows
传递给 Sheet.ForEachRow 以跳过空行。

**函数签名：**
```go
func SkipEmptyRows(flags *rowVisitorFlags)
```

**参数：**
- `flags` (*rowVisitorFlags): 行访问标志

#### SkipEmptyCells
传递给 Row.ForEachCell 以跳过空单元格。

**函数签名：**
```go
func SkipEmptyCells(flags *cellVisitorFlags)
```

**参数：**
- `flags` (*cellVisitorFlags): 单元格访问标志

## 数据类型

### CellType
单元格类型枚举。

**常量：**
- `CellTypeString`: 字符串类型
- `CellTypeStringFormula`: 字符串公式类型
- `CellTypeNumeric`: 数值类型
- `CellTypeBool`: 布尔类型
- `CellTypeInline`: 内联字符串类型
- `CellTypeError`: 错误类型
- `CellTypeDate`: 日期类型

## 示例

### 创建带样式的 Excel 文件

```go
package main

import (
    "github.com/Denghaoran7617/xlsx"
)

func main() {
    file := xlsx.NewFile()
    sheet, _ := file.AddSheet("Sheet1")
    
    // 创建样式
    style := xlsx.NewStyle()
    style.Font = *xlsx.NewFont(12, "Arial")
    style.Font.Bold = true
    style.Fill = *xlsx.NewFill("solid", "FF00FF00", "FF00FF00")
    
    // 创建行和单元格
    row := sheet.AddRow()
    cell := row.AddCell()
    cell.Value = "Hello"
    cell.SetStyle(style)
    
    file.Save("styled.xlsx")
}
```

### 读取复杂类型数据

```go
package main

import (
    "fmt"
    "github.com/Denghaoran7617/xlsx"
)

func main() {
    file, _ := xlsx.OpenFile("data.xlsx")
    sheet := file.Sheets[0]
    row := sheet.Row(0)
    
    type Data struct {
        Name     string              `xlsx:"0"`
        Tags     []string            `xlsx:"1"`
        Metadata map[string]interface{} `xlsx:"2"`
    }
    
    var data Data
    err := row.ReadStruct(&data)
    if err != nil {
        panic(err)
    }
    
    fmt.Printf("Name: %s\n", data.Name)
    fmt.Printf("Tags: %v\n", data.Tags)
    fmt.Printf("Metadata: %v\n", data.Metadata)
}
```

## 许可证

BSD 许可证

## 贡献

欢迎贡献！请确保：
- 所有现有测试通过
- 有测试覆盖您的更改
- 添加了文档字符串（至少是公共函数）
- XML 使用符合 ECMA-376 标准

## 相关链接

- [完整 API 文档](https://pkg.go.dev/github.com/Denghaoran7617/xlsx)
- [教程](https://github.com/Denghaoran7617/xlsx/blob/master/tutorial/tutorial-ch.md)

