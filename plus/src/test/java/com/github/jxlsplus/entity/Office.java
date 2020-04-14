package com.github.jxlsplus.entity;

import lombok.Builder;
import lombok.Data;

import java.io.Serializable;

/**
 * 办公室信息（jxls 测试专用）
 *
 * @author tzg
 */
@Data
@Builder
public class Office implements Serializable {

    private String officeCode;
    private String officeName;

}
