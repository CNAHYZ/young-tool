package com.antaiib.framework.excel.core.processor.sheet;



import com.antaiib.framework.excel.core.annotations.ResponseExcel;

import javax.servlet.http.HttpServletResponse;


/**
 * sheet 写出处理器
 * @author pig-mesh
 * @date 2023/09/08
 */
public interface SheetWriteProcessor {

    /**
     *     是否支持
     * @param obj   参数
     * @return boolean
     */
    boolean support(Object obj);
    
    /**
     * 校验
     *
     * @param responseExcel 注解
     */
    void check(ResponseExcel responseExcel);
    
    /**
     * 返回的对象
     *
     * @param o             obj
     * @param response      输出对象
     * @param responseExcel 注解
     */
    void export(Object o, HttpServletResponse response, ResponseExcel responseExcel);
    
    /**
     * 写成对象
     *
     * @param o             obj
     * @param response      输出对象
     * @param responseExcel 注解
     */
    void write(Object o, HttpServletResponse response, ResponseExcel responseExcel);
    
}
