#  MergeCells 指令说明

### 简介

MergeCells 指令用于合并单元格。

### Excel 标记语法

```
jx:mergeCells(
    lastCell="Merge cell ranges"
    [, cols="Number of columns combined"]
    [, rows="Number of rows combined"]
    [, minCols="Minimum number of columns to merge"]
    [, minRows="Minimum number of rows to merge"]
)
```

属性说明：

* lastCells 是指向指令区域最后一个单元格的任何指令的通用属性
* cols 合并列数
* rows 合并的行数
* minCols 合并的最小列数
* minRows 合并的最小行数

> 注意：该指令只能在尚未合并的单元格上使用。如果在合并单元格的区域内使用该指令，则会发生异常。

用例：

```
jx:mergeCells(cols="2" rows="list.size()" minCols="2" minRows="2" lastCell="A1")
```

> 假设 `list.size()` 为 `6` ，指令设置的单元格为 `A1`，则以上用例会转换成 **POI** 的实现，如下：
>
> ```java
> int startRow = 0, startCol = 0, 
> 	rows = 6, cols = 2,
> 	minRows = 2, minCols = 2
> 	endRow = startRow + Math.max(minRows, rows) - 1,
> 	endCol = startCol + Math.max(minCols, cols) - 1;
> CellRangeAddress region = new CellRangeAddress(startRow, endRow, startCol, endCol);
> // …… orther code
> ```

### 转换器支持说明

该指令当前只支持 **POI** 的实现。