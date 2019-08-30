# Excel 标记说明

**Jxls** 的 Excel 标记由三个部分组成：

* Bean 属性标记

* 区域定义标记

* 指令定义标记

**Jxls** 提供了 **XlsCommentAreaBuilder** 类，该类可以从 Excel 单元格注释中读取标记。
**XlsCommentAreaBuilder** 实现了通用的 **AreaBuilder** 接口。

```java
public interface AreaBuilder {
    List<Area> build();
}
```

它是一个简单的接口，只有一个方法，该方法返回一个区域对象列表。

因此，如果希望定义自己的标记，可以创建自己的 **AreaBuilder** 实现，并根据需要解析输入的 Excel 模板（或任何其他输入）。

### Bean 属性标记

**Jxls** 使用 [**Apache JEXL**](http://commons.apache.org/proper/commons-jexl/reference) 表达式语言进行处理

在将来的版本中，表达式语言引擎将是可配置的，因此如果需要，可以用任何其他表达式引擎替换 **JEXL** 。

**JEXL** 表达式语言语法请看[这里（**JEXL Syntax**）](http://commons.apache.org/proper/commons-jexl/reference/syntax.html)。

**Jxls** 希望 **JEXL** 表达式放在 XLS 模板文件中的 **${** **}** 中。

例如，单元格下的内容 **${department.chief.age} years** 将会告诉 **Jxls** 表达式解析器 **department.chief.age** ，
假设上下文对象 **Context** 中有一个可用的 **department** 变量对象，
如果这个对象可以通过 **department.getChief().getAge()** 表达式获取到值为 **35**，
那么 **Jxls** 在 **XlsArea** 处理期间，将在此单元格中填充 **35 years** 。

### 区域定义标记

**Jxls** 区域标记用于定义 **Jxls** 引擎可处理的根 **XlsArea** 。**XlsCommentAreaBuilder** 支持以下 Excel 单元格注释的语法作为区域定义：

```text
jx:area(lastCell="<LAST_CELL>")
```

这里的 **<LAST_CELL>** 定义矩形区域的右下角单元格。起始单元格的定义由放置该 Excel 注释的单元格决定。

因此，假设我们在单元格 **A1** 中有下一个注释jx:area(lastCell="G12")，根区域将被读取为 **A1:G12** 。

应该使用 **XlsCommentAreaBuilder** 从模板文件中读取所有区域。如以下代码片段，将所有区域读入 xlsAreaList，然后将第一个区域保存到 xlsArea 变量中

```java
AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
List<Area> xlsAreaList = areaBuilder.build();
Area xlsArea = xlsAreaList.get(0);
```

在大多数情况下，定义一个根 XlsArea 就足够了。

### 指令定义标记

指令应该在 XlsArea 中定义。**XlsCommentAreaBuilder** 接受以下作为 Excel 单元格注释创建的指令符号

```text
jx:<command_name>(attr1='val1' attr2='val2' ... attrN='valN' lastCell=<last_cell> areas=["<command_area1>", "<command_area2", ... "<command_areaN>"])
```

* **<command_name>** 是一个预先注册或手动在 **XlsCommentAreaBuilder** 中注册的命令名。目前，预注册了以下命令名：

    * each
    
    * if
    
    * image
    
    * grid

    自定义命令可以用 **XlsCommentAreaBuilder** 类的静态方法 **void addCommandMapping(String commandName, Class clazz)** 手动注册。

* **attr1**， **attr2**，...， **attrN** 是指令特定的属性。例如，**If-Command** 具有 **condition** 属性来设置条件表达式。

* **<last_cell>** 定义了命令体区域的右下角单元格。左上角的单元格由附加命令符号的单元格决定。

* **<command_ara1>**， **<command_area2>**，...**<command_areaN>** - XLS areas 作为参数传递给命令。

    例如，**If-Command** 希望定义以下区域

    * **ifArea** 是对 **If-command** 条件表达式计算为 **true** 时要输出的区域的引用

    * **elseArea** 是对 **If-command** 条件表达式计算为 **false** 时要输出的区域的引用(可选)

    因此，要为 **If-command** 定义 **area** 属性，它的 areas 属性如下：

    ```text
    areas=["A8:F8","A13:F13"]
    ```

在一个单元格注释中，您可以定义多个命令。例如，**Each** 和 **If** 命令定义如下所示：

```text
jx:each(items="department.staff", var="employee", lastCell="F8")
jx:if(condition="employee.payment <= 2000", lastCell="F8", areas=["A8:F8","A13:F13"])
```

> [原文链接](http://jxls.sourceforge.net/reference/excel_markup.html)
