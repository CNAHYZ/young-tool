package com.antaiib.framework.excel.core.head;

import lombok.Data;

import java.util.List;
import java.util.Set;


/**
 * @author pig-mesh
 * @date 2023/09/08
 */
@Data
public class HeadMeta {
    
    /**
     * <p>
     * 自定义头部信息
     * </p>
     * 实现类根据数据的class信息，定制Excel头<br/> 具体方法使用参考：https://www.yuque.com/easyexcel/doc/write#b4b9de00
     */
    private List<List<String>> head;
    
    /**
     * 忽略头对应字段名称
     */
    private Set<String> ignoreHeadFields;
    
}
