package com.github.jxlsplus.template;

import com.github.jxlsplus.command.JxlsPlusGridCommand;
import com.github.jxlsplus.transform.JxlsPlusPoiTransformer;
import lombok.extern.slf4j.Slf4j;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.template.SimpleExporter;
import org.jxls.transform.Transformer;

import java.io.*;
import java.util.List;

/**
 * 通用模板导出辅助类（参照 {@link SimpleExporter} 的实现）<br>
 * 注意：模板的文件格式与导出文件的格式要一致，默认格式为 .xls
 *
 * @author tzg
 */
@Slf4j
public class JxlsPlusExporter {

    private byte[] templateBytes;

    public JxlsPlusExporter() {
    }

    /**
     *
     * @param inputStream
     * @return
     */
    public JxlsPlusExporter registerGridTemplate(InputStream inputStream) {
        if (inputStream != null) {
            try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
                byte[] data = new byte[4096];
                int count;
                while ((count = inputStream.read(data)) != -1) {
                    os.write(data, 0, count);
                }
                templateBytes = os.toByteArray();
            } catch (IOException e) {
                String message = "Failed to read template file ";
                log.error(message);
                throw new JxlsException(message, e);
            }
        }
        return this;
    }

    /**
     * 用指定的通用模板导出数据文件（注意模板的文件格式与导出文件的格式要一致，默认格式为 .xls）
     *
     * @param headers
     * @param dataObjects
     * @param objectProps
     * @param outputStream
     */
    public void gridExport(Iterable<?> headers, Iterable<?> dataObjects, List<String> objectProps, OutputStream outputStream) {
        if (templateBytes == null) {
            InputStream is = SimpleExporter.class.getResourceAsStream(SimpleExporter.GRID_TEMPLATE_XLS);

            registerGridTemplate(is);
        }
        InputStream is = new ByteArrayInputStream(templateBytes);
        Transformer transformer = JxlsPlusPoiTransformer.createTransformer(is, outputStream);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
        Context context = JxlsPlusPoiTransformer.createInitialContext();
        context.putVar("headers", headers);
        context.putVar("data", dataObjects);
        JxlsPlusGridCommand gridCommand = (JxlsPlusGridCommand) xlsArea.getCommandDataList().get(0).getCommand();
        gridCommand.setRowObjectProps(objectProps);
        xlsArea.applyAt(new CellRef("Sheet1!A1"), context);
        try {
            transformer.write();
        } catch (IOException e) {
            final String msg = "Failed to write to output stream";
            log.error(msg, e);
            throw new JxlsException(msg, e);
        }
    }
}
