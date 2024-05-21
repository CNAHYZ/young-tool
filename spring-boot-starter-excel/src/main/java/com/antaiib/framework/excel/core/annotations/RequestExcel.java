package com.antaiib.framework.excel.core.annotations;

import com.antaiib.framework.excel.core.constants.ErrorExportTypeEnum;
import com.antaiib.framework.excel.core.listener.AbstractAnalysisEventListener;
import com.antaiib.framework.excel.core.listener.DefaultEasyExcelListener;
import com.antaiib.framework.excel.core.validator.IExcelDataChecker;

import java.lang.annotation.*;


/**
 * 导入excel
 *
 * @author pig-mesh
 * @date 2023/09/08
 */
@Documented
@Target({ElementType.PARAMETER})
@Retention(RetentionPolicy.RUNTIME)
public @interface RequestExcel {

    /**
     * 前端上传字段名称 file
     */
    String fileName() default "file";

    /**
     * 读取的监听器类
     *
     * @return readListener
     */
    Class<? extends AbstractAnalysisEventListener> readListener() default DefaultEasyExcelListener.class;

    /**
     * 是否跳过空行
     *
     * @return 默认跳过
     */
    boolean ignoreEmptyRow() default false;

    /**
     * 读取时使用的自定义检查器
     *
     * @return IExcelChecker
     */
    Class<? extends IExcelDataChecker> dataChecker() default IExcelDataChecker.class;

    /**
     * 是否开启分片导入
     */
    boolean enableSharding() default false;

    /**
     * 每次分片导入的大小，默认每隔1000条存储数据库，然后清理list ，方便内存回收
     */
    int shardingSize() default 1000;

    /**
     * 错误数据导出类型
     */
    ErrorExportTypeEnum errorExportType() default ErrorExportTypeEnum.ONLY_ERROR_DATA;


}
