# 自定义 **Command** （指令）

Jxls 允许根据自己的需求自定义指令。

### 怎么定义指令

编写自定义指令，必须要继承 **org.jxls.command.AbstractCommand** 抽象类，该类内容如下

```java
public abstract class AbstractCommand implements Command {
    private Logger logger = LoggerFactory.getLogger(AbstractCommand.class);
    List<Area> areaList = new ArrayList<Area>();
    private String shiftMode;

    @Override
    public Command addArea(Area area) {
        areaList.add(area);
        area.setParentCommand(this);
        return this;
    }

    @Override
    public void reset() {
        for (Area area : areaList) {
            area.reset();
        }
    }

    @Override
    public void setShiftMode(String mode) {
        if (mode != null) {
            if (mode.equalsIgnoreCase(Command.INNER_SHIFT_MODE) || mode.equalsIgnoreCase(Command.ADJACENT_SHIFT_MODE)) {
                shiftMode = mode;
            } else {
                logger.error("Cannot set cell shift mode to " + mode + " for command: " + getName());
            }
        }
    }

    @Override
    public String getShiftMode() {
        return shiftMode;
    }

    @Override
    public List<Area> getAreaList() {
        return areaList;
    }

    protected Transformer getTransformer() {
        if (areaList.isEmpty()) {
            return null;
        }
        return areaList.get(0).getTransformer();
    }

    protected TransformationConfig getTransformationConfig() {
        return getTransformer().getTransformationConfig();
    }
}
```

接下来我们实现一个分组的 **GroupRow** 指令，如下

```java
public class GroupRowCommand extends AbstractCommand {
    Area area;
    String collapseIf;

    @Override
    public String getName() {
        return "groupRow";
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Size resultSize = area.applyAt(cellRef, context);
        if( resultSize.equals(Size.ZERO_SIZE)) return resultSize;
        PoiTransformer transformer = (PoiTransformer) area.getTransformer();
        Workbook workbook = transformer.getWorkbook();
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        int startRow = cellRef.getRow();
        int endRow = cellRef.getRow() + resultSize.getHeight() - 1;
        sheet.groupRow(startRow, endRow);
        if( collapseIf != null && collapseIf.trim().length() > 0){
            boolean collapseFlag = Util.isConditionTrue(getTransformationConfig().getExpressionEvaluator(), collapseIf, context);
            sheet.setRowGroupCollapsed(startRow, collapseFlag);
        }
        return resultSize;
    }

    @Override
    public Command addArea(Area area) {
        super.addArea(area);
        this.area = area;
        return this;
    }

    public void setCollapseIf(String collapseIf) {
        this.collapseIf = collapseIf;
    }
}

```

### 怎么注册指令

使用 **XlsCommentAreaBuilder.addCommandMapping** 注册指令，如下：

```java
XlsCommentAreaBuilder.addCommandMapping("groupRow", GroupRowCommand.class);
```

> 注意：注册指令的时候，注册指令名称要与注册的指令中的 **getName()** 获取的值一致

### 怎么使用指令

使用方法可参照 [Each 指令](jxls_plus_each_command.md)

具体实现可查看 [jxls-demos](https://bitbucket.org/leonate/jxls-demo) 项目中的 **org.jxls.demo.UserCommandExcelMarkupDemo**
