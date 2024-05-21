package com.antaiib.framework.excel.core.config;

import com.alibaba.excel.converters.Converter;
import com.antaiib.framework.excel.core.aop.ResponseExcelReturnValueHandler;
import com.antaiib.framework.excel.core.enhance.WriterBuilderEnhancer;
import com.antaiib.framework.excel.core.head.I18nHeaderCellWriteHandler;
import com.antaiib.framework.excel.core.processor.sheet.ManySheetWriteProcessor;
import com.antaiib.framework.excel.core.processor.sheet.SheetWriteProcessor;
import com.antaiib.framework.excel.core.processor.sheet.SingleSheetWriteProcessor;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.boot.autoconfigure.condition.ConditionalOnBean;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.context.MessageSource;
import org.springframework.context.annotation.Bean;

import javax.annotation.Resource;
import java.util.List;


/**
 * @author pig-mesh
 * @date 2023/09/08
 */
@RequiredArgsConstructor
public class ExcelHandlerConfiguration {
    
    private final ExcelConfigProperties configProperties;
    
    private final ObjectProvider<List<Converter<?>>> converterProvider;
    @Resource
    private WriterBuilderEnhancer writerBuilderEnhancer;
    

    
    /**
     * 单sheet 写入处理器
     */
    @Bean
    @ConditionalOnMissingBean
    public SingleSheetWriteProcessor singleSheetWriteHandler() {
        return new SingleSheetWriteProcessor(configProperties, converterProvider, writerBuilderEnhancer);
    }
    
    /**
     * 多sheet 写入处理器
     */
    @Bean
    @ConditionalOnMissingBean
    public ManySheetWriteProcessor manySheetWriteHandler() {
        return new ManySheetWriteProcessor(configProperties, converterProvider, writerBuilderEnhancer);
    }
    
    /**
     * 返回Excel文件的 response 处理器
     *
     * @param sheetWriteHandlerList 页签写入处理器集合
     * @return ResponseExcelReturnValueHandler
     */
    @Bean
    @ConditionalOnMissingBean
    public ResponseExcelReturnValueHandler responseExcelReturnValueHandler(
            List<SheetWriteProcessor> sheetWriteHandlerList) {
        return new ResponseExcelReturnValueHandler(sheetWriteHandlerList);
    }
    
    /**
     * excel 头的国际化处理器
     *
     * @param messageSource 国际化源
     */
    @Bean
    @ConditionalOnBean(MessageSource.class)
    @ConditionalOnMissingBean
    public I18nHeaderCellWriteHandler i18nHeaderCellWriteHandler(MessageSource messageSource) {
        return new I18nHeaderCellWriteHandler(messageSource);
    }
    
}
