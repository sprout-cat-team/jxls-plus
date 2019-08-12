package com.tang.jxlsplus.command;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.command.GridCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.util.UtilWrapper;

import java.util.*;
import java.util.regex.Pattern;

import static org.apache.commons.lang3.StringUtils.EMPTY;
import static org.jxls.command.GridCommand.DATA_VAR;
import static org.jxls.command.GridCommand.HEADER_VAR;

/**
 * 扩展 {@link GridCommand} 指令。<br>
 * 目前采取拷贝原本进行修改的方式，因为指令中大部分变量和方法是所有的无法重写。<br>
 * 读取修改的源码版本为 v2.6.0 <br>
 * 增加“var”和“propsVar”变量。 <br>
 * 1、<b>var</b> ：允许 grid 指令在指令中指定元素的变量名称，默认为 <b>rowData</b> ； <br>
 * 2、<b>propsVar</b> : 允许 grid 指令通过变量获取元素属性变量名，多个时，以分号（;）隔开；
 *
 * @author tzg
 */
@Slf4j
public class JxlsPlusGridCommand extends AbstractCommand {

    public static final String ROW_VAR = "rowData";

    /**
     * Name of a context variable containing a collection of headers
     */
    private String headers;
    /**
     * Name of a context variable containing a collection of data objects for body
     */
    private String data;
    /**
     * 行数据的参数名
     */
    private String var = ROW_VAR;
    /**
     * Comma separated list of object property names for each grid row
     */
    private String props;
    /**
     * 从变量中获取属性值
     */
    private String propsVar;
    /**
     * Comma separated list of format type cells, e.g. formatCells="Double:E1, Date:F1"
     */
    private String formatCells;
    private Map<String, String> cellStyleMap = new HashMap<>();
    private List<String> rowObjectProps = new ArrayList<>();
    private Area headerArea;
    private Area bodyArea;

    protected UtilWrapper util = new UtilWrapper();
    protected ExpressionEvaluator expressionEvaluator;

    public JxlsPlusGridCommand() {
    }

    public JxlsPlusGridCommand(String headers, String data) {
        this.headers = headers;
        this.data = data;
    }

    public JxlsPlusGridCommand(String headers, String data, String props, Area headerArea, Area bodyArea) {
        this(headers, data, headerArea, bodyArea);
        this.props = props;
    }

    public JxlsPlusGridCommand(String headers, String data, Area headerArea, Area bodyArea) {
        this.headers = headers;
        this.data = data;
        this.headerArea = headerArea;
        this.bodyArea = bodyArea;
        addArea(headerArea);
        addArea(bodyArea);
    }

    @Override
    public String getName() {
        return GridCommand.COMMAND_NAME;
    }

    public String getHeaders() {
        return headers;
    }

    public void setHeaders(String headers) {
        this.headers = headers;
    }

    public String getData() {
        return data;
    }

    public void setData(String data) {
        this.data = data;
    }

    public String getVar() {
        return var;
    }

    public void setVar(String var) {
        this.var = var;
    }

    public String getProps() {
        return props;
    }

    /**
     * 设置属性（多个值时以分号隔开）
     *
     * @param props
     */
    public void setProps(String props) {
        this.props = props;
        if (props != null) {
            rowObjectProps = Arrays.asList(StringUtils.split(props.replaceAll("\\s+", EMPTY), ";")); // Remove whitespace and split into List.
        }
    }

    public String getPropsVar() {
        return propsVar;
    }

    public void setPropsVar(String propsVar) {
        this.propsVar = propsVar;
    }

    public String getFormatCells() {
        return formatCells;
    }

    /**
     * @param formatCells Comma-separated list of format type cells, e.g. formatCells="Double:E1, Date:F1"
     */
    public void setFormatCells(String formatCells) {
        this.formatCells = formatCells;
        if (StringUtils.isNotBlank(formatCells)) {
            String[] cellStyleList = StringUtils.split(formatCells, ",");
            try {
                for (String cellStyleString : cellStyleList) {
                    if (StringUtils.containsNone(cellStyleString, ":")) {
                        continue;
                    }
                    String[] styleCell = StringUtils.split(cellStyleString, ":");
                    cellStyleMap.put(styleCell[0].trim(), styleCell[1].trim());
                }
            } catch (Exception e) {
                log.error("Failed to parse formatCells attribute");
            }
        }
    }

    @Override
    public Command addArea(Area area) {
        if (getAreaList().size() >= 2) {
            throw new IllegalArgumentException("Cannot add any more areas to GridCommand. You can add only 1 area as a 'header' and 1 area as a 'body'.");
        }
        if (getAreaList().isEmpty()) {
            headerArea = area;
        } else {
            bodyArea = area;
        }
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        expressionEvaluator = getTransformationConfig().getExpressionEvaluator();
        Size headerAreaSize = processHeaders(cellRef, context);
        CellRef bodyCellRef = new CellRef(cellRef.getSheetName(), cellRef.getRow() + headerAreaSize.getHeight(), cellRef.getCol());
        Size bodyAreaSize = processBody(bodyCellRef, context);
        int gridHeight = headerAreaSize.getHeight() + bodyAreaSize.getHeight();
        int gridWidth = Math.max(headerAreaSize.getWidth(), bodyAreaSize.getWidth());
        return new Size(gridWidth, gridHeight);
    }

    /**
     * 表头的处理
     *
     * @param cellRef
     * @param context
     * @return
     */
    private Size processHeaders(CellRef cellRef, Context context) {
        if (headerArea == null || headers == null) {
            return Size.ZERO_SIZE;
        }
        if (StringUtils.isNotBlank(propsVar)) {
            // 从变量中获取属性值
            String propStr = (String) expressionEvaluator.evaluate(propsVar, context.toMap());
            setProps(propStr);
        }
//        Iterable<?> headers = util.transformToIterableObject(expressionEvaluator, this.headers, context);
        Object headerObj = expressionEvaluator.evaluate(this.headers, context.toMap());
        if (headerObj == null) {
            return Size.ZERO_SIZE;
        }
        Collection headerCollection;
        if (headerObj.getClass().isArray()) {
            headerCollection = Arrays.asList((Object[]) headerObj);
        } else if (headerObj instanceof Collection) {
            headerCollection = (Collection) headerObj;
        } else {
            throw new JxlsException(this.headers + " expression is not a collection or a array");
        }
        CellRef currentCell = cellRef;
        int width = 0;
        int height = 0;
        for (Object header : headerCollection) {
            if (header.getClass().isArray() || header instanceof Iterable) {
                // TODO 复杂表头的处理
                Iterable cellCollection;
                if (header.getClass().isArray()) {
                    cellCollection = Arrays.asList((Object[]) header);
                } else {
                    cellCollection = (Iterable) header;
                }
                for (Object cellObject : cellCollection) {
                    context.putVar(HEADER_VAR, cellObject);
                    Size size = headerArea.applyAt(currentCell, context);
                    currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                }

                currentCell = new CellRef(cellRef.getSheetName(), currentCell.getRow() + height, cellRef.getCol());
            } else {
                context.putVar(HEADER_VAR, header);
                Size size = headerArea.applyAt(currentCell, context);
                currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                width += size.getWidth();
                height = Math.max(height, size.getHeight());
            }
        }
        context.removeVar(HEADER_VAR);

        return new Size(width, height);
    }

    /**
     * 数据内容的处理
     *
     * @param cellRef
     * @param context
     * @return
     */
    private Size processBody(final CellRef cellRef, Context context) {
        if (bodyArea == null || data == null) {
            return Size.ZERO_SIZE;
        }
        Iterable<?> dataCollection = util.transformToIterableObject(expressionEvaluator, this.data, context);

        CellRef currentCell = cellRef;
        int totalWidth = 0;
        int totalHeight = 0;
        Context.Config config = context.getConfig();
        boolean oldIgnoreSourceCellStyle = config.isIgnoreSourceCellStyle();
        config.setIgnoreSourceCellStyle(true);
        Map<String, String> oldStyleCellMap = config.getCellStyleMap();
        config.setCellStyleMap(this.cellStyleMap);
        // TODO possible error: content of DATA_VAR is not saved & restored
        for (Object rowObject : dataCollection) {
            context.putVar(var, rowObject);
            if (rowObject.getClass().isArray() || rowObject instanceof Iterable) {
                Iterable<?> cellCollection;
                if (rowObject.getClass().isArray()) {
                    cellCollection = Arrays.asList((Object[]) rowObject);
                } else {
                    cellCollection = (Iterable<?>) rowObject;
                }
                int width = 0;
                int height = 0;
                for (Object cellObject : cellCollection) {
                    context.putVar(DATA_VAR, cellObject);
                    Size size = bodyArea.applyAt(currentCell, context);
                    currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                    width += size.getWidth();
                    height = Math.max(height, size.getHeight());
                }
                totalWidth = Math.max(width, totalWidth);
                totalHeight = totalHeight + height;
                currentCell = new CellRef(cellRef.getSheetName(), currentCell.getRow() + height, cellRef.getCol());
            } else {
                if (rowObjectProps.isEmpty()) {
                    throw new IllegalArgumentException("Got a non-collection object type for a Grid row but object properties list is empty");
                }
                int width = 0;
                int height = 0;
                for (String prop : rowObjectProps) {
                    try {
                        Object value = checkRowData(prop) ?
                                // 支持通过字段映射获取对应的值
                                expressionEvaluator.evaluate(prop, context.toMap()) :
                                PropertyUtils.getProperty(rowObject, prop);
                        context.putVar(DATA_VAR, value);
                        Size size = bodyArea.applyAt(currentCell, context);
                        currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                        width += size.getWidth();
                        height = Math.max(height, size.getHeight());
                    } catch (Exception e) {
                        String message = "Failed to evaluate property " + prop + " of row object of class " + rowObject.getClass().getName();
                        log.error(message, e);
                        throw new IllegalStateException(message, e);
                    }
                }
                totalWidth = Math.max(width, totalWidth);
                totalHeight = totalHeight + height;
                currentCell = new CellRef(cellRef.getSheetName(), currentCell.getRow() + height, cellRef.getCol());
            }
        }
        context.removeVar(DATA_VAR);
        context.removeVar(var);
        config.setIgnoreSourceCellStyle(oldIgnoreSourceCellStyle);
        config.setCellStyleMap(oldStyleCellMap);
        return new Size(totalWidth, totalHeight);
    }

    /**
     * 检查属性字符串是否包含 rowData.*
     *
     * @param prop
     * @return
     */
    protected boolean checkRowData(String prop) {
        if (StringUtils.isAnyBlank(prop, var)) {
            return false;
        } else {
            return Pattern.compile(var.concat("\\.+\\w")).matcher(prop).find();
        }
    }

}
