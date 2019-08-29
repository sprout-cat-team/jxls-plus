package com.tang.jxlsplus;

import com.tang.jxlsplus.entity.Employee;
import com.tang.jxlsplus.entity.Office;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.SystemUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
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
import java.text.ParseException;
import java.time.LocalDate;
import java.util.*;

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
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\jp_eachDemo.xlsx", currentUser),
                context
        );

        log.debug("===================================== 分割线 ==================================================");

        jxlsHelper(
                getClass().getResourceAsStream(template),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\eachDemo.xlsx", currentUser),
                context
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
        context.putVar("headers", headers);
        context.putVar("cellProps", StringUtils.join(cellProps, ";"));

        String template = "/grid_template.xlsx";
        JxlsPlusUtils.processTemplate(
                getClass().getResourceAsStream(template),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\jp_gridDemo.xlsx", currentUser),
                context
        );

        JxlsPlusUtils.processGridTemplate(
                Arrays.asList(headers),
                employees,
                cellProps,
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\jp_gridDemo2.xls", currentUser)
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

    private String currentUser = "administrator";
    private List<Employee> employees = new ArrayList<>();
    private Context context;

    /**
     * 生成用例数据
     */
    @Before
    public void generateSampleData() throws ParseException {
        currentUser = System.getProperty("user.name");

        employees.add(
                Employee.builder()
                        .empId(UUID.randomUUID().toString())
                        .name(currentUser + "_超长内容_超长内容_超长内容_超长内容_\r\n超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容")
                        .age(30)
                        .birthDate(LocalDate.of(1998, 1, 17))
                        .hiredate(DateFormatUtils.ISO_DATE_FORMAT.parse("2002-01-18"))
                        .bonus(15000D)
                        .mobile("15013459999")
                        .email("tang_guo_168@163.com")
                        .build()
        );

        Date currentDate = new Date();
        for (int i = 0; i < 100; i++) {
            employees.add(
                    Employee.builder()
                            .empId(UUID.randomUUID().toString())
                            .name("Employee " + (i + 1))
                            .age(18 + i % 3)
                            .mobile("15013459999")
                            .email("tang_guo_" + i + "@163.com")
                            .birthDate(LocalDate.of(2001 - i % 2, 1, 1))
                            .hiredate(currentDate)
                            .bonus(1000D * (i % 2 + 1))
                            .offices(initOfices(i % 2 == 0, 2))
                            .build()
            );
        }

        Map<String, Object> params = new HashMap<String, Object>() {
            {
                put("employees", employees);
                put("nowdate", new Date());
            }
        };


        context = new Context(params);
    }

    protected List<Office> initOfices(boolean isEven, final int size) {
        if (size < 1) {
            return Collections.emptyList();
        }
        List<Office> list = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            Office.OfficeBuilder officeBuilder = Office.builder();
            if (isEven) {
                officeBuilder.officeCode("ok" + i)
                        .officeName("OK 办公室" + i);
            } else {
                officeBuilder.officeCode("no" + i)
                        .officeName("NO 办公室" + i);
            }
            list.add(officeBuilder.build());
        }
        return list;
    }

}
