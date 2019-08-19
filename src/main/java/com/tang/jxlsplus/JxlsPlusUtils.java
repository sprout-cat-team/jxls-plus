package com.tang.jxlsplus;

import com.tang.jxlsplus.command.JxlsPlusEachCommand;
import com.tang.jxlsplus.command.JxlsPlusGridCommand;
import com.tang.jxlsplus.formula.JxlsPlusFormulaProcessor;
import com.tang.jxlsplus.transform.JxlsPlusPoiTransformer;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.GridCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.formula.FormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.util.List;

/**
 * JxlsPlus 工具类 <br>
 * 注意：默认会有注入 stringUtils ({@link StringUtils}) 和 dateUtils ({@link DateFormatUtils}) 两个工具类 <br>
 * 用法为：${dateUtils.formatUTC(varName,"yyyy-MM-dd")}
 *
 * @author tzg
 */
@Slf4j
public class JxlsPlusUtils {

    static {
        // 自定义指令的注册
        XlsCommentAreaBuilder.addCommandMapping(JxlsPlusEachCommand.COMMAND_NAME, JxlsPlusEachCommand.class);
        XlsCommentAreaBuilder.addCommandMapping(GridCommand.COMMAND_NAME, JxlsPlusGridCommand.class);
    }

    /**
     * 字符串处理工具 key
     */
    public final static String UTILS_STR = "stringUtils";
    /**
     * 日期处理工具 key
     */
    public final static String UTILS_DATE = "dateUtils";

    /**
     * 根据输入的模板导出 excel 文件（基础实现）<br>
     * 参照 {@link JxlsHelper#createTransformer} 的实现
     *
     * @param inputStream  输入的模板文件流
     * @param outputStream 输出的文件流
     * @param context      参数配置项
     */
    public static void processTemplate(@NonNull InputStream inputStream, @NonNull OutputStream outputStream, Context context) {
        if (log.isDebugEnabled()) {
            log.debug("start export excel file");
        }
        if (null == context) {
            context = JxlsPlusPoiTransformer.createInitialContext();
        }
        if (!context.toMap().containsKey(UTILS_STR)) {
            // 添加字符串处理工具方法
            context.putVar(UTILS_STR, new StringUtils());
        }
        if (!context.toMap().containsKey(UTILS_DATE)) {
            // 添加日期处理工具方法
            context.putVar(UTILS_DATE, new DateFormatUtils());
        }
        try {
//        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
//        Transformer transformer = jxlsHelper.createTransformer(inputStream, outputStream);
            Transformer transformer = JxlsPlusPoiTransformer.createTransformer(inputStream, outputStream);

            AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
            areaBuilder.setTransformer(transformer);
            List<Area> xlsAreaList = areaBuilder.build();
            FormulaProcessor formulaProcessor = new JxlsPlusFormulaProcessor();
            for (Area xlsArea : xlsAreaList) {
                xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
                xlsArea.setFormulaProcessor(formulaProcessor);
                xlsArea.processFormulas();
            }
            transformer.write();
        } catch (InvalidFormatException e) {
            throw new IllegalArgumentException("文件格式不正确", e);
        } catch (IOException e) {
            throw new IllegalArgumentException("文件读写异常", e);
        } finally {
            if (null != inputStream) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (null != outputStream) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 根据输入的模板导出 excel 文件
     *
     * @param templateFile 模板文件
     * @param outputStream 输出的文件流
     * @param context
     */
    public static void processTemplate(@NonNull File templateFile, OutputStream outputStream, Context context) {
        if (!templateFile.exists()) {
            throw new IllegalArgumentException("This template file not exist");
        }
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(templateFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        processTemplate(inputStream, outputStream, context);
    }

    /**
     * 根据输入的模板导出 excel 文件
     *
     * @param templateFile 模板文件
     * @param outFile      输出文件
     * @param context
     */
    public static void processTemplate(File templateFile, File outFile, Context context) {
        processTemplate(templateFile, newOutputStream(outFile), context);
    }

    /**
     * 创建文件的输出流（判断文件的文件夹不存在时创建新的）
     *
     * @param outFile
     * @return
     */
    public static OutputStream newOutputStream(File outFile) {
        //检测目录是否存在，不存在则创建
        if (!outFile.getParentFile().exists()) {
            // 如果文件所在的目录不存在，则创建目录
            if (!outFile.getParentFile().mkdir()) {
                if (log.isDebugEnabled()) {
                    log.error("This file directory[{}] make failure", outFile.getParentFile().getAbsolutePath());
                }
                throw new IllegalArgumentException("This file directory make failure");
            }
        }
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(outFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return outputStream;
    }


    /**
     * 根据输入的模板导出 excel 文件
     *
     * @param templatePath 模板文件路径
     * @param outputStream
     * @param context
     */
    public static void processTemplate(@NonNull String templatePath, OutputStream outputStream, Context context) {
        processTemplate(new File(templatePath), outputStream, context);
    }

    /**
     * 根据输入的模板导出 excel 文件
     *
     * @param inputStream
     * @param outputPath
     * @param context
     */
    public static void processTemplate(InputStream inputStream, @NonNull String outputPath, Context context) {
        File outFile = new File(outputPath);
        processTemplate(inputStream, newOutputStream(outFile), context);
    }

}
