package com.github.jxlsplus;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.junit.Before;
import org.junit.Test;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.EachCommand;
import org.jxls.command.GridCommand;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * jxls plus 测试用例（用例中的输出路径是针对 win 的，如果有其他系统请自行修改）
 * <br> TODO 还需要添加原始指令的测试用例
 *
 * @author tzg
 */
@Slf4j
public class JxlsPlusTest {

    @Test
    public void eachCmd() throws IOException {
        String template = "/each_template.xlsx";
        JxlsPlusUtils.processTemplate(
                getClass().getResourceAsStream(template),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\jp_eachDemo.xlsx", testData.getCurrentUser()),
                testData.getContext()
        );

        log.debug("===================================== 分割线 ==================================================");

        jxlsHelper(
                getClass().getResourceAsStream(template),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\eachDemo.xlsx", testData.getCurrentUser()),
                testData.getContext()
        );
    }

    @Test
    public void gridCmd() throws IOException {
        String[] headers = new String[]{"序号", "姓名", "年龄", "电子邮箱", "入职日期"};
        List<String> cellProps = new ArrayList<String>() {{
            add("rowData_index+1");
            add("name");
            add("age");
            add("email");
            add("hiredate");
        }};
        testData.getContext().putVar("headers", headers);
        testData.getContext().putVar("cellProps", StringUtils.join(cellProps, ";"));

        String template = "/grid_template.xlsx";
        JxlsPlusUtils.processTemplate(
                getClass().getResourceAsStream(template),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\jp_gridDemo.xlsx", testData.getCurrentUser()),
                testData.getContext()
        );

        JxlsPlusUtils.processGridTemplate(
                Arrays.asList(headers),
                testData.getEmployees(),
                cellProps,
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\jp_gridDemo2.xls", testData.getCurrentUser())
        );

//        log.debug("===================================== 分割线 ==================================================");
//
//        jxlsHelper(
//                getClass().getResourceAsStream(template),
//                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\gridDemo.xlsx", currentUser),
//                context
//        );
    }

    //======================== fun          =====================================

    /**
     * jxls 原始导出方法
     *
     * @param templateStream
     * @param ouputPath
     * @param context
     * @throws IOException
     */
    protected void jxlsHelper(InputStream templateStream, String ouputPath, Context context) throws IOException {
        // 重置指令
        XlsCommentAreaBuilder.addCommandMapping(EachCommand.COMMAND_NAME, EachCommand.class);
        XlsCommentAreaBuilder.addCommandMapping(GridCommand.COMMAND_NAME, GridCommand.class);
        JxlsHelper.getInstance().processTemplate(
                templateStream,
                JxlsPlusUtils.newOutputStream(new File(ouputPath)),
                context
        );
    }

    //======================== 数据准备阶段 =====================================

    private JxlsPlusTestData testData;

    /**
     * 生成用例数据
     */
    @Before
    public void generateSampleData() {
        testData = JxlsPlusTestData.init();
    }


}
