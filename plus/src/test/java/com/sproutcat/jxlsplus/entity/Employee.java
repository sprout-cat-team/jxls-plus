package com.sproutcat.jxlsplus.entity;

import lombok.Builder;
import lombok.Data;

import java.io.Serializable;
import java.time.LocalDate;
import java.util.Date;
import java.util.List;

/**
 * 用户信息（jxls 测试专用）
 *
 * @author tzg
 */
@Data
@Builder
public class Employee implements Serializable {

    private String empId;
    private String name;
    private Integer age;
    private LocalDate birthDate;
    private String mobile;
    private String email;
    /**
     * 奖金
     */
    private Double bonus;
    /**
     * 入职时间
     */
    private Date hiredate;
    private List<Office> offices;

}
