package com.github.sproutcat.jxlsplus;

import com.github.sproutcat.jxlsplus.command.JxlsPlusEachCommand;
import com.github.sproutcat.jxlsplus.command.JxlsPlusGridCommand;
import com.github.sproutcat.jxlsplus.template.JxlsPlusExporter;
import com.github.sproutcat.jxlsplus.transform.JxlsPlusPoiTransformer;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.GridCommand;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
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
     * 根据输入的模板导出 excel 文件（基础实现）<br>
     * 参照 {@link JxlsHelper#createTransformer} 的实现
     *
     * @param templateStream 输入的模板文件流
     * @param targetStream   输出的文件流
     * @param context        参数配置项（可以是数据变量或工具类）
     */
    public static void processTemplate(@NonNull InputStream templateStream, @NonNull OutputStream targetStream, Context context) {
        if (log.isDebugEnabled()) {
            log.debug("start export excel file");
        }
        if (null == context) {
            context = JxlsPlusPoiTransformer.createInitialContext();
        } else {
            JxlsPlusPoiTransformer.checkContext(context);
        }

        JxlsHelper jxlsHelper = JxlsHelper.getInstance();

        try {
            // 创建自定义的转换器
            Transformer transformer = JxlsPlusPoiTransformer.createTransformer(templateStream, targetStream);
            // 设置自定义公式处理器
//            jxlsHelper.setFormulaProcessor(new JxlsPlusFormulaProcessor());
            jxlsHelper.processTemplate(context, transformer);
        } catch (IOException e) {
            throw new JxlsException("Failed to write to output stream", e);
        } finally {
            try {
                templateStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                targetStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 根据输入的模板导出 excel 文件
     *
     * @param templateFile 模板文件
     * @param targetStream 输出的文件流
     * @param context      参数配置项（可以是数据变量或工具类）
     */
    public static void processTemplate(@NonNull File templateFile, OutputStream targetStream, Context context) {
        if (!templateFile.exists()) {
            throw new IllegalArgumentException("This template file not exist");
        }
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(templateFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        processTemplate(inputStream, targetStream, context);
    }

    /**
     * 根据输入的模板导出 excel 文件
     *
     * @param templateFile 模板文件
     * @param outFile      输出文件
     * @param context      参数配置项（可以是数据变量或工具类）
     */
    public static void processTemplate(File templateFile, File outFile, Context context) {
        processTemplate(templateFile, newOutputStream(outFile), context);
    }

    /**
     * 创建文件的输出流（判断文件的文件夹不存在时创建新的）
     *
     * @param outFile 输出文件
     * @return
     */
    public static OutputStream newOutputStream(@NonNull File outFile) {
        //检测目录是否存在，不存在则创建
        if (!outFile.getParentFile().exists()) {
            // 如果文件所在的目录不存在，则创建目录
            if (!outFile.getParentFile().mkdir()) {
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
     * @param targetStream
     * @param context      参数配置项（可以是数据变量或工具类）
     */
    public static void processTemplate(@NonNull String templatePath, OutputStream targetStream, Context context) {
        processTemplate(new File(templatePath), targetStream, context);
    }

    /**
     * 根据输入的模板导出 excel 文件
     *
     * @param templateStream 输入的模板文件流
     * @param outputPath     输出文件路径
     * @param context        参数配置项（可以是数据变量或工具类）
     */
    public static void processTemplate(InputStream templateStream, @NonNull String outputPath, Context context) {
        File outFile = new File(outputPath);
        processTemplate(templateStream, newOutputStream(outFile), context);
    }

    /**
     * 指定通用模板导出 excel 文件 <br>
     * {@link JxlsPlusExporter} 的实现
     *
     * @param targetStream   目标输出流
     * @param headers        表头记录
     * @param dataObjects    数据记录
     * @param objectProps    数据属性
     * @param templateStream 模板输入流
     */
    public static void processGridTemplate(InputStream templateStream, Iterable<?> headers, Iterable<?> dataObjects,
                                           List<String> objectProps, OutputStream targetStream) {
        new JxlsPlusExporter().registerGridTemplate(templateStream)
                .gridExport(headers, dataObjects, objectProps, targetStream);
    }

    /**
     * 通用模板导出 excel 文件<br>
     * {@link JxlsPlusExporter} 的实现
     *
     * @param headers      表头记录
     * @param dataObjects  数据记录
     * @param cellProps    数据属性
     * @param targetStream 目标输出流
     */
    public static void processGridTemplate(Iterable<?> headers, Iterable<?> dataObjects,
                                           List<String> cellProps, OutputStream targetStream) {
        new JxlsPlusExporter().gridExport(headers, dataObjects, cellProps, targetStream);
    }

    /**
     * 通用模板导出 excel 文件
     *
     * @param headers     表头记录
     * @param dataObjects 数据记录
     * @param cellProps   数据属性
     * @param targetPath  目标文件路径
     */
    public static void processGridTemplate(Iterable<?> headers, Iterable<?> dataObjects,
                                           List<String> cellProps, String targetPath) {
        OutputStream targetStream = newOutputStream(new File(targetPath));
        processGridTemplate(headers, dataObjects, cellProps, targetStream);
    }

}
