# Grid 指令说明

**Grid** 指令对于生成具有头和数据行区域的动态网格非常有用。

表头作为字符串的集合传递，数据行作为对象或列表的集合传递。

### 指令的使用

**Grid** 指令有以下属性：

* **headers** - 包含标题集合的上下文变量的名称。

* **data** - 包含数据集合的上下文变量的名称。

* **props** - 用于指定每个网格行的对象属性，多个时以逗号分隔开（仅当每个网格行是对象时才有效）。

* **formatCells** - 数据类型与单元格格式的映射，多个时以逗号分隔开，例如 **formatCells="Double:E1, Date:F1"** 。

* **headerArea** - 表头的原始 xls 区域。

* **bodyArea** - 数据的原始 xls 区域。

* **lastCell** - 是指向命令区域最后一个单元格的任何命令的通用属性。

> **data** 变量可以是以下类型<br>
> 1、**`Collection<Collection<Object>>`** - 这里，每个内部集合包含对应行的单元格值。<br>
> 2、**`Collection<Object>`** - 这里，每个集合项都是一个对象，其中包含对应行的数据。
> 在这种情况下，必须指定 **props** 属性来定义应该使用哪些对象属性来设置特定单元格的数据。<br>
> <br>
> **jxls-plus** 项目对 **Grid** 指令的扩展：<br>
> 1、**var** ：允许 **Grid** 指令在指令中指定元素的变量名称，默认为 **rowData** ； <br>
> 2、**propsVar** : 允许 **Grid** 指令通过变量获取元素属性变量名，多个时以分号（;）隔开；<br>
> 3、**headers** 属性接收字符串数组对象。

当通过 **header** 集合进行迭代时，**Grid** 指令将每个 header 放入 **header** 键下的上下文中。
在对数据按行迭代的过程中，每个单元格的数据项都放在 **cell** 键下的上下文中。

因此，在 Excel 模板 **Grid** 指令只需要两个单元格，一个是表头单元格，一个是数据行单元格。表头单元格可以定义为：

```text
 ${header}
```

数据行单元格可以定义为

```text
${cell}
```

**Grid** 指令使用样例请查看[动态网络样例](http://jxls.sourceforge.net/samples/dynamic_grid.html)

### 转换器支持说明

请注意，**Grid** 指令目前仅在 **POI** 转换器中受支持，因此如果您想使用 **Grid** 指令，则必须使用 **POI**。

> [原文地址](http://jxls.sourceforge.net/reference/grid_command.html)
