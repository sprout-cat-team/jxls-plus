# **Jxls** 主要概念

**Jxls** 基于以下主要概念

* **XlsArea** - 矩形区域

* **Command** - 指令

* **Transformer** - 转换器

接下来让我们详细讨论这些概念。

### XlsArea - 矩形区域

**XlsArea** 表示 Excel 文件中的矩形区域。它可以使用单元格范围定义，也可以通过指定区域的起始单元格和大小（列和行数）来定义。

**XlsArea** 包含指定范围内的所有 **Excel** 单元格。

每个 **XlsArea** 可能都有一组与其相关联的命令，这些命令将在 Jxls 引擎处理区域期间执行。
**XlsArea** 可能有嵌套在其中的子区域。每个子区域也是 **XlsArea** ，具有自己的命令，并且可能有自己的子区域。

可以使用以下方法定义 **XlsArea** ：

* 在 Excel 模板中使用 **jxls** 特定的标记语法。**Jxls** 用它的 **XlsCommentAreaBuilder** 提供了一个默认的标记。如果需要，可以定义自定义标记。

* 通过使用 **XML** 配置。**Jxls** 提供 **XmlAreaBuilder** 类作为 **XML** 标记的默认实现。

* 通过使用 **Jxls** Java API

您可以在这里找到更多的细节和使用[示例](http://jxls.sourceforge.net/reference/xls_area.html)

### Command - 指令

**Command** （指令）表示单个或多个 **XlsArea** 上的转换操作。

对应的 Java 接口如下所示：

```java
public interface Command {
    String getName();
    List<Area> getAreaList();
    Command addArea(Area area);
    Size applyAt(CellRef cellRef, Context context);
    void reset();
}
```

**Command** 的主要方法是 **Size applyAt(CellRef cellRef, Context context)** 。
该方法使 **Command** 可以使用 **context** 传递变量数据，在单元格 **cellRef** 上进行操作。
**context** 用一个 Map 实现，用于将变量数据传递给 **Command** 。
该方法返回一个新维度的矩形区域 **Size** 对象。

目前，**Jxls** 提供了以下内置命令：

* [**Each-Command**](jxls_plus_each_command.md) - 需要对一个集合进行处理时可使用该指令。

* **If-Command** - 该指令用于处理判断表达式。

* **MergeCells-command** - 该指令用于合并单元格。

* **Image-Command** - 该指令用于渲染图片。

如果默认提供的指令不能满足您的需求，你还可以根据需求[自定义指令](jxls_custom_command.md)。

使用 **context** 对象将数据传递给 **Command** ，在 XLS 模板中使用 **context** 对象的引用键获取对应的值，也可根据情况在指令中设置需要的数据。

### Transformer - 转换器

**Transformer** 接口允许 **XlsArea** 独立于任何特定的与 Excel 交互的实现。
这意味着通过提供不同的 **Transformer** 接口实现，我们可以使用不同的底层 Java-to-Excel 库。

实现接口如下：

```java
public interface Transformer {
    void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeight);

    void setFormula(CellRef cellRef, String formulaString);

    Set<CellData> getFormulaCells();

    CellData getCellData(CellRef cellRef);

    List<CellRef> getTargetCellRef(CellRef cellRef);

    void resetTargetCellRefs();

    void resetArea(AreaRef areaRef);

    void clearCell(CellRef cellRef);

    List<CellData> getCommentedCells();

    void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType);

    void write() throws IOException;

    TransformationConfig getTransformationConfig();

    void setTransformationConfig(TransformationConfig transformationConfig);

    boolean deleteSheet(String sheetName);

    void setHidden(String sheetName, boolean hidden);

    void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum);
}
```

虽然它看起来有很多方法，但是其中很多方法已经在基础抽象类 **AbstractTransformer** 中实现，如果需要支持新的 Java-to-Excel 实现，可以继承它们。

目前，**Jxls** 提供了接口的两种实现

* **PoiTransformer** - 基于 [**Apache POI**](https://poi.apache.org/) 的实现

* **JexcelTransformer** - 基于 [**Java Excel API**](http://jexcelapi.sourceforge.net/) 的实现

注意：**Jxls-plus** 是针对 **PoiTransformer** 进行扩展

### [原文地址](http://jxls.sourceforge.net/reference/main_concepts.html)
