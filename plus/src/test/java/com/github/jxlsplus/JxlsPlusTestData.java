package com.github.jxlsplus;

import com.github.jxlsplus.entity.Employee;
import com.github.jxlsplus.entity.Office;
import lombok.Builder;
import lombok.Data;
import org.apache.commons.lang3.time.DateUtils;
import org.jxls.common.Context;

import java.text.ParseException;
import java.time.LocalDate;
import java.util.*;

/**
 * 测试数据准备
 */
@Data
@Builder
public class JxlsPlusTestData {

    private String currentUser;
    private List<Employee> employees;
    private Context context;

    /**
     * 初始化生成用例数据
     */
    public static JxlsPlusTestData init() {
        String currentUser = System.getProperty("user.name");
        JxlsPlusTestDataBuilder dataBuilder = JxlsPlusTestData.builder()
                .currentUser(currentUser);


        List<Employee> employees = new ArrayList<>();

        try {
            employees.add(
                    Employee.builder()
                            .empId(UUID.randomUUID().toString())
                            .name(currentUser + "_超长内容_超长内容_超长内容_超长内容_\r\n超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容_超长内容")
                            .age(30)
                            .birthDate(LocalDate.of(1998, 1, 17))
                            .hiredate(DateUtils.parseDate("2002-01-18", "yyyy-MM-dd"))
                            .bonus(15000D)
                            .mobile("15013459999")
                            .email("tang_guo_168@163.com")
                            .build()
            );
        } catch (ParseException e) {
            e.printStackTrace();
        }

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

        dataBuilder.employees(employees).context(new Context(params));
        return dataBuilder.build();
    }

    protected static List<Office> initOfices(boolean isEven, final int size) {
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
