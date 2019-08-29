# XLS Area

### 简介


### 构建 XLS Area


### 使用 Excel 标记构建 XLS Area

您可以在 Excel 模板中使用特殊的标记来构造 XLS Area 。标记应该放在该区域第一个单元格的 Excel 注释中。标记写法如下：

```text
jx:area(lastCell = "<AREA_LAST_CELL>")
```

其中 **<AREA_LAST_CELL>** 是定义区域的最后一个单元格。


### 使用 XML 配置来构建 XLS Area


### 使用 Java API 来构建 XLS Area

To create an XLS Area with Java API you may use one of the **XlsArea** class constructors. The following constructors are available

```java
public XlsArea(AreaRef areaRef, Transformer transformer);

public XlsArea(String areaRef, Transformer transformer);

public XlsArea(CellRef startCell, CellRef endCell, Transformer transformer);

public XlsArea(CellRef startCellRef, Size size, List<CommandData> commandDataList, Transformer transformer);

public XlsArea(CellRef startCellRef, Size size);

public XlsArea(CellRef startCellRef, Size size, Transformer transformer);
```

To build a top level area you have to provide a Transformer instance so that the area can use it for transformation.

And you have to define the area cells using cell range as a string or alternatively by creating CellRef cell reference object and set the area Size.

Here is a snippet of code to construct a set of nested template XLS areas with commands

```java
// create Transformer instance
// ...
// Create a top level area
XlsArea xlsArea = new XlsArea("Template!A1:G15", transformer);
// Create 'department' are
XlsArea departmentArea = new XlsArea("Template!A2:G13", transformer);
// create 'EachCommand' to iterate through departments
EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
// create an area for employee 'each' command
XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
// create an area for 'if' command
XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
// create 'if' command with the specified areas
IfCommand ifCommand = new IfCommand("employee.payment <= 2000",
        ifArea,
        new XlsArea("Template!A9:F9", transformer));
// adding 'if' command instance to employee area
employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
// create employee 'each' command and add it to department area
Command employeeEachCommand = new EachCommand( "employee", "department.staff", employeeArea);
departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
// add department 'each' command to top-level area
xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
```

### [原文地址](http://jxls.sourceforge.net/reference/xls_area.html)
