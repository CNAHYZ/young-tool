package com.antaiib.framework.excel.core.enhance;

import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.antaiib.framework.excel.core.annotations.ResponseExcel;
import com.antaiib.framework.excel.core.head.HeadGenerator;

import javax.servlet.http.HttpServletResponse;


/**
 * ExcelWriterBuilder 增强
 * @author pig-mesh
 * @date 2023/09/08
 */
public interface WriterBuilderEnhancer {
    
    /**
     * ExcelWriterBuilder 增强
     *
     * @param writerBuilder ExcelWriterBuilder
     * @param response      HttpServletResponse
     * @param responseExcel ResponseExcel
     * @param templatePath  模板地址
     * @return ExcelWriterBuilder
     */
    ExcelWriterBuilder enhanceExcel(ExcelWriterBuilder writerBuilder, HttpServletResponse response,
                                    ResponseExcel responseExcel, String templatePath);
    
    /**
     * ExcelWriterSheetBuilder 增强
     *
     * @param writerSheetBuilder ExcelWriterSheetBuilder
     * @param sheetNo            sheet角标
     * @param sheetName          sheet名，有模板时为空
     * @param dataClass          当前写入的数据所属类
     * @param template           模板文件
     * @param headEnhancerClass  当前指定的自定义头处理器
     * @return ExcelWriterSheetBuilder
     */
    ExcelWriterSheetBuilder enhanceSheet(ExcelWriterSheetBuilder writerSheetBuilder, Integer sheetNo, String sheetName,
            Class<?> dataClass, String template, Class<? extends HeadGenerator> headEnhancerClass);
    
}
