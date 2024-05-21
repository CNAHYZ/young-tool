package com.antaiib.framework.excel.core.annotations;



import com.antaiib.framework.excel.core.head.HeadGenerator;

import java.lang.annotation.*;


/**
 * @author pig-mesh
 * @date 2023/09/08
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Sheet {
    
    int sheetNo() default -1;
    
    /**
     * sheet name
     */
    String sheetName();
    
    /**
     * 包含字段
     */
    String[] includes() default {};
    
    /**
     * 排除字段
     */
    String[] excludes() default {};
    
    /**
     * 头生成器
     */
    Class<? extends HeadGenerator> headGenerateClass() default HeadGenerator.class;
    
}
