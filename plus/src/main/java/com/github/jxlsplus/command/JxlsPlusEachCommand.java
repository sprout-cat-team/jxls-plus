package com.github.jxlsplus.command;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.jxls.area.Area;
import org.jxls.command.*;
import org.jxls.common.*;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.util.JxlsHelper;
import org.jxls.util.UtilWrapper;

import java.util.*;

/**
 * Description: 扩展 {@link EachCommand} 指令 <br>
 * 1、增加下标变量“var_index”。如var="item"，获取下标方法：${item_index}
 * 2、增加分组下标变量“var_g_index”
 * 3、扩展对空数组的处理
 *
 * @author tzg
 */
@Slf4j
public class JxlsPlusEachCommand extends EachCommand {

    public JxlsPlusEachCommand() {
    }

    public JxlsPlusEachCommand(String var, String items, Direction direction) {
        super(var, items, direction);
    }

    public JxlsPlusEachCommand(String items, Area area) {
        super(items, area);
        this.area = area;
    }

    public JxlsPlusEachCommand(String var, String items, Area area) {
        super(var, items, area);
        this.area = area;
    }

    public JxlsPlusEachCommand(String var, String items, Area area, Direction direction) {
        super(var, items, area, direction);
        this.area = area;
    }

    public JxlsPlusEachCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
        super(var, items, area, cellRefGenerator);
        this.area = area;
    }

    protected static final String GROUP_DATA_KEY = "_group";

    protected UtilWrapper util = new UtilWrapper();
    protected Area area;


    /**
     * 空数据时，是否添加空白的区域，默认添加区域
     */
    private boolean blankArea = true;

    public boolean isBlankArea() {
        return blankArea;
    }

    public void setBlankArea(boolean blankArea) {
        this.blankArea = blankArea;
    }

    public void setBlankArea(String blankArea) {
        if (StringUtils.isBlank(blankArea)) {
            this.blankArea = false;
        } else {
            this.blankArea = "1".equals(blankArea) || "true".equalsIgnoreCase(blankArea);
        }
    }

    @Override
    public Command addArea(Area area) {
        this.area = area;
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Iterable<?> itemsCollection;
        Object collectionObject = getTransformationConfig().getExpressionEvaluator().evaluate(getItems(), context.toMap());
        if (collectionObject instanceof Iterable) {
            itemsCollection = (Iterable<?>) collectionObject;
        } else {
            log.warn("Failed to evaluate collection expression {}", getItems());
            itemsCollection = Collections.emptyList();
        }
        Size size;
        if (getGroupBy() == null || getGroupBy().length() == 0) {
            size = processCollection(context, itemsCollection, cellRef, getVar(), false);
        } else {
            Collection<GroupData> groupedData = util.groupIterable(itemsCollection, getGroupBy(), getGroupOrder());
            String groupVar = getVar() != null ? getVar() : GROUP_DATA_KEY;
            size = processCollection(context, groupedData, cellRef, groupVar, true);
        }
        if (getDirection() == Direction.DOWN) {
            getTransformer().adjustTableSize(cellRef, size);
        }
        return size;
    }

    /**
     * 数据集合的处理,针对（{@link EachCommand#processCollection(Context, Iterable, CellRef, String)}）方法的重写
     * @param context
     * @param itemsCollection
     * @param cellRef
     * @param varName
     * @param isGroup
     * @return
     */
    protected Size processCollection(Context context, Iterable<?> itemsCollection, CellRef cellRef, String varName, boolean isGroup) {
        int index = 0, dataIndex = 0, gIndex = 0, newWidth = 0, newHeight = 0;

        CellRefGenerator cellRefGenerator = this.getCellRefGenerator();
        if (cellRefGenerator == null && getMultisheet() != null) {
            List<String> sheetNameList = extractSheetNameList(context);
            cellRefGenerator = sheetNameList == null
                    ? new DynamicSheetNameGenerator(getMultisheet(), cellRef, getTransformationConfig().getExpressionEvaluator())
                    : new SheetNameGenerator(sheetNameList, cellRef);
        }

        ExpressionEvaluator selectEvaluator = null;
        if (getSelect() != null) {
            selectEvaluator = JxlsHelper.getInstance().createExpressionEvaluator(getSelect());
        }

        String varNameIndex = varName.concat("_index"), groupIndex = varName.concat("_g_index");
        if (!isGroup) {
            Object g = context.getVar(groupIndex);
            if (null != g) {
                gIndex = ((int) context.getVar(groupIndex)) + 1;
            }
        }
        CellRef currentCell = cellRef;
        Object currentVarObject = context.getVar(varName);
        if (null != itemsCollection && itemsCollection.iterator().hasNext()) {
            for (Object obj : itemsCollection) {
                context.putVar(varName, obj);
                context.putVar(varNameIndex, dataIndex);
                if (!isGroup) {
                    context.putVar(groupIndex, gIndex);
                    gIndex++;
                }
                if (selectEvaluator != null && !util.isConditionTrue(selectEvaluator, context)) {
                    context.removeVar(varName);
                    context.removeVar(varNameIndex);
                    continue;
                }
                if (cellRefGenerator != null) {
                    currentCell = cellRefGenerator.generateCellRef(index++, context);
                }
                if (currentCell == null) {
                    break;
                }
                Size size = area.applyAt(currentCell, context);
                if (!Size.ZERO_SIZE.equals(size)) {
                    // 空行不计数
                    dataIndex++;
                }
                if (cellRefGenerator != null) {
                    newWidth = Math.max(newWidth, size.getWidth());
                    newHeight = Math.max(newHeight, size.getHeight());
                } else if (getDirection() == Direction.DOWN) {
                    currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow() + size.getHeight(), currentCell.getCol());
                    newWidth = Math.max(newWidth, size.getWidth());
                    newHeight += size.getHeight();
                } else { // RIGHT
                    currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                    newWidth += size.getWidth();
                    newHeight = Math.max(newHeight, size.getHeight());
                }
            }
            if (currentVarObject != null) {
                context.putVar(varName, currentVarObject);
                context.putVar(varNameIndex, dataIndex);
                context.putVar(groupIndex, gIndex);
            } else {
                context.removeVar(varName);
                context.removeVar(varNameIndex);
                if (isGroup) {
                    context.removeVar(groupIndex);
                }
            }
            return new Size(newWidth, newHeight);
        } else if (blankArea) {
            // 保证没有数据时，至少有一个区域
            return area.applyAt(currentCell, context);
        } else {
            return Size.ZERO_SIZE;
        }
    }

    /**
     * 抽取工作簿名集合（{@link EachCommand#extractSheetNameList}）
     *
     * @param context
     * @return
     */
    protected List<String> extractSheetNameList(Context context) {
        try {
            Object sheetnames = context.getVar(getMultisheet());
            if (sheetnames == null) {
                return null;
            } else if (sheetnames instanceof Collection) {
                return new ArrayList<>((Collection<String>) sheetnames);
            } else if (sheetnames.getClass().isArray()) {
                return Arrays.asList((String[]) sheetnames);
            }
        } catch (Exception e) {
            throw new JxlsException("Failed to get sheet names from " + getMultisheet(), e);
        }
        throw new JxlsException("The sheet names var '" + getMultisheet() + "' must be of type List<String>.");
    }

}
