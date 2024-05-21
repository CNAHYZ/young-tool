package com.antaiib.framework.excel.core.processor;

import java.lang.reflect.Method;


/**
 * @author pig-mesh
 * @date 2023/09/08
 */
public interface NameProcessor {
    
    /**
     * 解析名称
     *
     * @param args   拦截器对象
     * @param key    表达式
     */
    String doDetermineName(Object[] args, Method method, String key);
    
}
