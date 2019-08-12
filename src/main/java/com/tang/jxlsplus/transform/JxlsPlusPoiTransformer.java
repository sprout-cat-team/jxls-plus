package com.tang.jxlsplus.transform;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.jxls.common.*;
import org.jxls.transform.poi.PoiCellData;
import org.jxls.transform.poi.PoiTransformer;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * 对 {@link PoiTransformer} 的扩展，主要针对高度自适应的处理
 *
 * @author tzg
 */
@Slf4j
public class JxlsPlusPoiTransformer extends PoiTransformer {

    /**
     * No streaming
     *
     * @param workbook
     */
    private JxlsPlusPoiTransformer(Workbook workbook) {
        this(workbook, false);
    }

    public JxlsPlusPoiTransformer(Workbook workbook, boolean streaming) {
        super(workbook, streaming);
    }

    public JxlsPlusPoiTransformer(Workbook workbook, boolean streaming, int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        super(workbook, streaming, rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
    }

    /**
     * 重写 {@link PoiTransformer#transformCell(CellRef, CellRef, Context, boolean, CellData, Sheet, Row)} 方法，针对高度自适应进行处理
     *
     * @param srcCellRef
     * @param targetCellRef
     * @param context
     * @param updateRowHeightFlag
     * @param cellData
     * @param destSheet
     * @param destRow
     */
    @Override
    protected void transformCell(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeightFlag, CellData cellData, Sheet destSheet, Row destRow) {
        SheetData sheetData = sheetMap.get(srcCellRef.getSheetName());
        if (!isIgnoreColumnProps()) {
            destSheet.setColumnWidth(targetCellRef.getCol(), sheetData.getColumnWidth(srcCellRef.getCol()));
        }
        if (updateRowHeightFlag && !isIgnoreRowProps()) {
            Row row = destSheet.getRow(targetCellRef.getRow());
            short height = row.getHeight(),
                    srcHeight = (short) sheetData.getRowData(srcCellRef.getRow()).getHeight();
            CellStyle style = row.getRowStyle();
            boolean wrapText = null == style ? false : style.getWrapText();
            if (wrapText && height != (short) -1) {
                // 根据内容自适应高度的处理
                row.setHeight((short) -1);
            } else {
                // 避免自定义的高度被恢复为原始的高度
                if (height < srcHeight) {
                    row.setHeight(srcHeight);
                }
            }
        }
        org.apache.poi.ss.usermodel.Cell destCell = destRow.getCell(targetCellRef.getCol());
        if (destCell == null) {
            destCell = destRow.createCell(targetCellRef.getCol());
        }
        try {
            destCell.setCellType(CellType.BLANK);
            ((PoiCellData) cellData).writeToCell(destCell, context, this);
            copyMergedRegions(cellData, targetCellRef);
        } catch (Exception e) {
            log.error("Failed to write a cell with {} and context keys {}", cellData, context.toMap().keySet(), e);
        }
    }

    /**
     * 获取 row data
     *
     * @param sheetName
     * @param rowNum
     * @return
     */
    public RowData getRowData(String sheetName, int rowNum) {
        return sheetMap.get(sheetName).getRowData(rowNum);
    }

    /**
     * 实例化转换器（必须要实现，不可用 {@link PoiTransformer#createTransformer}）
     *
     * @param is
     * @param os
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static JxlsPlusPoiTransformer createTransformer(InputStream is, OutputStream os) throws IOException, InvalidFormatException {
        JxlsPlusPoiTransformer transformer = createTransformer(is);
        transformer.setOutputStream(os);
        transformer.setInputStream(is);
        return transformer;
    }

    public static JxlsPlusPoiTransformer createTransformer(InputStream is) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(is);
        return createTransformer(workbook);
    }

    public static JxlsPlusPoiTransformer createTransformer(Workbook workbook) {
        return new JxlsPlusPoiTransformer(workbook);
    }

}
