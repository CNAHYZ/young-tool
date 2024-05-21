package com.antaiib.framework.excel.core.config;


import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;


/**
 * @author pig-mesh
 * @date 2023/09/08
 */
@Data
@ConfigurationProperties(prefix = ExcelConfigProperties.PREFIX)
public class ExcelConfigProperties {
    
    static final String PREFIX = "excel";
    
    /**
     * 模板路径
     */
    private String templatePath = "excel";
    
}
