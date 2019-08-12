package com.tang.jxlsplus.entity;

import lombok.Builder;
import lombok.Data;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 用户信息（jxls 测试专用）
 * @author tzg
 */
@Data
@Builder
public class Employee implements Serializable {

    private String empId;
    private String name;
    private Integer age;
    private String mobile;
    private String email;
    private Date lastLoginDate;
    private LocalDateTime loginFailureDate;
    private BigDecimal loginCount;
    private List<Office> offices;

}
