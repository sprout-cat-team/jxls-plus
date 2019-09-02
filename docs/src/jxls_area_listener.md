# 区域侦听器（ AreaListener ）

区域侦听器（ AreaListener ）可用于执行响应区域转换事件的附加区域处理。例如，您可能希望根据数据突出显示某些行或单元格。

### AreaListener 接口

**AreaListener** 接口如下

```java
public interface AreaListener {
    void beforeApplyAtCell(CellRef cellRef, Context context);
    void afterApplyAtCell(CellRef cellRef, Context context);
    void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context);
    void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context);
}
```

当转换相应 xls 区域中的单元格时，将调用相应的方法。每个侦听器方法都获取对正在转换的单元格的单元格引用（ **CellRef cellRef** ）和上下文（ **Context context** ）。
转换侦听器方法还获取目标单元格引用（ **CellRef targetCell** ）。

* **beforeApplyAtCell()** 方法在单元格处理开始之前调用

* **afterApplyAtCell()** 方法在单元格处理完成后调用

* **beforeTransformCell()** 方法在单元格即将由转换器转换时调用

* **afterTransformCell()** 方法在单元格被转换器转换后调用

请参阅[区域侦听器示例](http://jxls.sourceforge.net/samples/area_listener.html)以查看它的实际操作。

> [原文地址](http://jxls.sourceforge.net/reference/area_listener.html)
