package com.antaiib.framework.excel.core.config;

import com.antaiib.framework.excel.core.enhance.DefaultWriterBuilderEnhancer;
import com.antaiib.framework.excel.core.enhance.WriterBuilderEnhancer;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * @author pig-mesh
 * @version 1.0
 * @date 2023/8/30 13:23
 */
@Configuration
public class WriterBuilderEnhancerConfig {

    @Bean
    @ConditionalOnMissingBean
    public WriterBuilderEnhancer writerBuilderEnhancer() {
        return new DefaultWriterBuilderEnhancer();
    }
}
