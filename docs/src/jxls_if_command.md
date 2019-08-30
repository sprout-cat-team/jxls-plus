# If 指令

### 简介

If 指令是一个条件指令，它根据指令的 *test* 属性中指定的条件输出区域。

### 指令属性

如果指令具有以下属性

* **condition** 是要测试的条件表达式。

* **ifArea** 是当此指令条件计算为 true 时要输出的区域的引用。

* **elseArea** 是当此指令条件计算为 false 时要输出的区域的引用。

* **lastCell** 是指向指令区域的最后一个单元格的任何指令的公共属性。

> **ifArea** 和 **condition** 属性是必要属性。

### 指令编译

与任何 Jxls 指令一样，您可以使用 Java API 或 Excel 标记或 XML 配置来定义 **If** 指令。

###### Java API 使用

下面是在 BitBucket 示例的 Jxls 示例中创建 if-Command 的示例。

```java
// ...
// creating 'if' and 'else' areas
XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
XlsArea elseArea = new XlsArea("Template!A9:F9", transformer);
// creating 'if' command
IfCommand ifCommand = new IfCommand("employee.payment <= 2000", ifArea, elseArea);
```

##### Excel 标记

若要使用 Excel 标记创建 if-Command，应在指令主体区域的起始单元格注释中使用以下语法

```text
jx:if(condition="employee.payment <= 2000", lastCell="F9", areas=["A9:F9","A18:F18"])
```

这里，**lastCell** 属性定义了 **If** 指令作用区域的最后一个单元格。

##### XML 标记

若要使用XML配置创建 **If** 指令，请使用以下标记

```xml
<area ref="Template!A9:F9">
    <if condition="employee.payment &lt;= 2000" ref="Template!A9:F9">
        <area ref="Template!A18:F18"/>
        <area ref="Template!A9:F9"/>
    </if>
</area>
```

这里，**ref** 属性定义了与 **If** 指令相关联的区域。

> [原文地址](http://jxls.sourceforge.net/reference/if_command.html)
