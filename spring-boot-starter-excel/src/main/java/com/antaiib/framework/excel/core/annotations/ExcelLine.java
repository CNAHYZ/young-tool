package com.antaiib.framework.excel.core.annotations;

import java.lang.annotation.*;

/**
 * 用于导入时获取excel 行号
 *
 * @author pig-mesh
 * @date 2023/08/30
 */
@Documented
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelLine {

}
