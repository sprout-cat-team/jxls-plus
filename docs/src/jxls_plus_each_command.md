# Each 指令说明

### 简介

**Each** 指令用于遍历集合并克隆指令 **XLS area** 。它是一个类似于 Java 的 **for** 操作符。

##### 指令属性

**Each** 有如下属性：

* _**var**_ 是 **Jxls** 上下文中用于在迭代时放置每个新集合项的变量的名称

* _**items**_ 是包含要迭代的集合的上下文变量的名称

* _**area**_ 是 **Each** 指令作用范围的 **XLS Area** 的引用

* _**direction**_ 是 **Direction** 枚举的值，它包含有 **DOWN** （向下）或 **RIGHT** （向右）的值，指示如何按行或按列循环指令体。默认值为 **DOWN**。

* _**select**_ 是一个表达式选择器，用于在迭代期间过滤掉集合项。

* _**groupBy**_ 是进行分组的属性

* _**groupOrder**_ 表示对分组的排序（‘ **desc** ’或‘ **asc** ’）

* _**cellRefGenerator**_ 是用于创建目标单元格引用的自定义策略

* _**multisheet**_ 是上下文变量的名称，其中包含要输出集合的工作表名称列表

* _**lastCell**_ 是指向指令区域最后一个单元格的任何指令的通用属性

> 其中 **var** 和 **items** 是必要的属性，而其他属性为可选属性。

有关使用 **cellRefGenerator** 和 **multisheet** 属性的更多信息，请检查[例子](http://jxls.sourceforge.net/reference/multi_sheets.html)。

### 指令构建

与任何 **Jxls** 指令一样，您可以使用 Java API 或 Excel 标记或 XML 配置来定义 **Each** 指令。

##### Java API 的使用

下面是创建 **Each** 指令的例子（可在 BitBucket 上查看 [Jxls 的示例](https://bitbucket.org/leonate/jxls-demo)）

```java
// creating a transformer and departments collection
...
// creating department XlsArea
XlsArea departmentArea = new XlsArea("Template!A2:G13", transformer);
// creating Each Command to iterate departments collection and attach to it "departmentArea"
EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
```

##### Excel 标记的使用

要使用 Excel 标记创建 **Each** 指令，您应该在注释中为指令主体区域的起始单元格使用特定的语法

```text
jx:each(items="employees" var="employee" lastCell="D4")
```

因此，我们使用的属性都在 **jx:each** 指令的括号中，并以空格分隔。**lastCell** 属性定义指令 **XlsArea** 作用范围的最后一个单元格。

##### XML 标记的使用

要使用 XML 配置创建 **Each** 指令，可以使用以下标记：

```xml
<each items="employees" var="employee" ref="Template!A4:D4">
    <area ref="Template!A4:D4"/>
</each>
```

上面的 **ref** 属性定义了与 **Each** 指令相关联的区域。内部 **area** 元素定义了 Each 指令的作用范围。通常它们是一样的。

### 怎么控制循环的方向

默认情况下，Each 指令的 **direction** 属性都被设置为 **DOWN** ，这意味着指令主体将向下循环 Excel 的行。

如果需要按列复制区域，应该将 **direction** 属性设置为 **RIGHT** 。

使用 Java API 可以这样做：

```java
//... creating EachCommand to iterate departments
// setting its direction to RIGHT
departmentEachCommand.setDirection(EachCommand.Direction.RIGHT);
```

### 怎么对数据进行分组

**Each** 指令支持通过 **groupBy** 属性对数据进行分组。通过 **groupOrder** 属性设置分组后数据的顺序，可以是 **desc** 或 **asc** 。

使用 Excel 标记可以参照一下配置：

```text
jx:each(items="employees" var="myGroup" groupBy="name" groupOrder="asc" lastCell="D6")
```

在本例中，Each 指令对数据进行分组后，可以使用 **myGroup** 变量引用从上下文中中访问的该比例的值。

可以使用 **myGroup.item** 引用当前组项。所以要获取员工姓名的使用

```text
${myGroup.item.name}
```

分组中的所有项都可以通过组的 **items** 属性访问，例如：

```text
jx:each(items="myGroup.items" var="employee" lastCell="D6")
```

您还可以完全忽略 **var** 属性，在本例中，默认的分组变量名称将是 **_group** 。

有关示例，请查看[分组示例](http://jxls.sourceforge.net/samples/grouping_example.html)

> [原文地址](http://jxls.sourceforge.net/reference/each_command.html)
