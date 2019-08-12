package com.tang.jxlsplus;

import com.tang.jxlsplus.entity.Employee;
import com.tang.jxlsplus.entity.Office;
import org.junit.Before;
import org.junit.Test;
import org.jxls.common.Context;

import java.util.*;

/**
 * jxls plus 测试用例（用例中的输出路径是针对 win 的，如果有其他系统请自行修改）
 *
 * @author tzg
 */
public class JxlsPlusTest {

    @Test
    public void gridCmd() {
        context.putVar("headers", new String[]{"Name", "Age", "Email", "LastLoginDate"});
        context.putVar("cellProps", "rowData.name;rowData.age;rowData.email;rowData.lastLoginDate");
        JxlsPlusUtils.export(
                getClass().getResourceAsStream("/grid_template.xlsx"),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\gridDemo.xlsx", currentUser),
                context
        );
    }

    @Test
    public void eachCmd() {
        JxlsPlusUtils.export(
                getClass().getResourceAsStream("/each_template.xlsx"),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\eachDemo.xlsx", currentUser),
                context
        );
    }

    @Test
    public void eachIfCmd() {
        JxlsPlusUtils.export(
                getClass().getResourceAsStream("/each_if_template.xlsx"),
                String.format("C:\\Users\\%s\\Desktop\\jxls-plus\\eachIfDemo.xlsx", currentUser),
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
    public void generateSampleData() {
        currentUser = System.getProperty("user.name");

        employees.add(
                Employee.builder()
                        .empId(UUID.randomUUID().toString())
                        .name(currentUser + "_超长内容_超长内容_超长内容_超长内容_\r\n超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容")
                        .age(30)
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
                            .age(18)
                            .mobile("15013459999")
                            .email("tang_guo_" + i + "@163.com")
                            .lastLoginDate(currentDate)
                            .offices(initOfices(i%2 == 0, 2))
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

    public List<Office> initOfices(boolean isEven, final int size) {
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
