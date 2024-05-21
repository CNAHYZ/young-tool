package com.antaiib.framework.excel.core.config;

import com.antaiib.framework.excel.core.aop.DynamicNameAspect;
import com.antaiib.framework.excel.core.aop.RequestExcelArgumentResolver;
import com.antaiib.framework.excel.core.aop.ResponseExcelReturnValueHandler;
import com.antaiib.framework.excel.core.processor.NameProcessor;
import com.antaiib.framework.excel.core.processor.NameSpelExpressionProcessor;
import lombok.RequiredArgsConstructor;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Import;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.servlet.mvc.method.annotation.RequestMappingHandlerAdapter;

import javax.annotation.PostConstruct;
import java.util.ArrayList;
import java.util.List;


/**
 * 配置初始化
 * @author pig-mesh
 * @date 2023/09/08
 */
@RequiredArgsConstructor
@Import(ExcelHandlerConfiguration.class)
@EnableConfigurationProperties(ExcelConfigProperties.class)
@Configuration
public class ResponseExcelAutoConfiguration {


    private final RequestMappingHandlerAdapter requestMappingHandlerAdapter;
    
    private final ResponseExcelReturnValueHandler responseExcelReturnValueHandler;
    
    /**
     *  解析处理器
     *
     * @return NameProcessor excel名称解析器
     */
    @Bean
    @ConditionalOnMissingBean
    public NameProcessor nameProcessor() {
        return new NameSpelExpressionProcessor();
    }
    
    /**
     * Excel名称解析处理切面
     *
     * @param nameProcessor  解析处理器
     * @return DynamicNameAspect
     */
    @Bean
    @ConditionalOnMissingBean
    public DynamicNameAspect dynamicNameAspect(NameProcessor nameProcessor) {
        return new DynamicNameAspect(nameProcessor);
    }
    
    /**
     * 追加 Excel返回值处理器 到 springmvc 中
     */
    @PostConstruct
    public void setReturnValueHandlers() {
        List<HandlerMethodReturnValueHandler> returnValueHandlers = requestMappingHandlerAdapter.getReturnValueHandlers();
        
        List<HandlerMethodReturnValueHandler> newHandlers = new ArrayList<>();
        newHandlers.add(responseExcelReturnValueHandler);
        assert returnValueHandlers != null;
        newHandlers.addAll(returnValueHandlers);
        requestMappingHandlerAdapter.setReturnValueHandlers(newHandlers);
    }
    
    /**
     * 追加 Excel 请求处理器 到 springmvc 中
     */
    @PostConstruct
    public void setRequestExcelArgumentResolver() {
        List<HandlerMethodArgumentResolver> argumentResolvers = requestMappingHandlerAdapter.getArgumentResolvers();
        List<HandlerMethodArgumentResolver> resolverList = new ArrayList<>();
        resolverList.add(new RequestExcelArgumentResolver());
        assert argumentResolvers != null;
        resolverList.addAll(argumentResolvers);
        requestMappingHandlerAdapter.setArgumentResolvers(resolverList);
    }
    
}
