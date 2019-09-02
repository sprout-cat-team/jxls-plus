# UpdateCell 指令说明

**UpdateCell** 指令允许您为特定的单元自定义处理。

### 指令属性

**UpdateCell** 指令包含有以下属性：

* **updater** 是包含 **CellDataUpdater** 实现的上下文中键的名称。

* **lastCell** 是指向指令区域最后一个单元格的任何指令的通用属性。

**CellDataUpdater** 接口如下

```java
public interface CellDataUpdater {
    void updateCellData(CellData cellData, CellRef targetCell, Context context);
}
```

在转换 **UpdateCell** 指令区域之前，将调用 **updateCellData** 方法，该方法传递当前单元格数据（ **CellData** ）、目标单元格索引（ **CellRef** ）和当前上下文（ **Context** ）。

实现可以更新 CellData 来设置单元格所需的值。

例如，下面是一个更新总公式的类

```java
class TotalCellUpdater implements CellDataUpdater{
    public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
        if(cellData.isFormulaCell() && cellData.getFormula().equals("SUM(E2)")){
            String resultFormula = String.format("SUM(E2:E%d)", targetCell.getRow());
            cellData.setEvaluationResult(resultFormula);
        }
    }
}
```

关键行是 **cellData.setEvaluationResult(resultFormula)** ，它使用目标公式更新单元格数据。

例如，在无法使用标准 **Jxls** 公式处理功能的 **SXSSF** 处理中，这可能很有用。

有关 **UpdateCell** 命令使用的更多示例，请参见 [jxls-demo](https://bitbucket.org/leonate/jxls-demo/) 仓库中的 SxssfDemo 示例。

### Excel 标记的使用

在 Excel 模板单元格注释中创建 **UpdateCell** 指令，指令语法如下

```text
jx:updateCell(lastCell="E4"  updater="totalCellUpdater")
```

**lastCell** 属性定义指令区域的最后一个单元格。

**updater** 属性被设置为 **totalCellUpdater** 。在处理之前，必须将 **totalCellUpdater** 放到上下文中

```java
Context context = new Context();
context.putVar("totalCellUpdater", new TotalCellUpdater());
```
