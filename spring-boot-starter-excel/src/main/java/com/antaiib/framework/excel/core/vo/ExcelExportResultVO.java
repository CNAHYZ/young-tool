package com.antaiib.framework.excel.core.vo;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * excel导出结果VO类
 *
 * @author yz
 * @since 2024/5/7 9:06
 */
@Data
@AllArgsConstructor
public class ExcelExportResultVO<T> {
    /**
     * 单选映射 key:列号 value:下拉选项
     */
    Map<Integer, List<String>> singleSelectorMap;

    /**
     * 数据
     */
    List<T> data;

    /**
     * 实际泛型类型，由于data为空时无法获取到实际类型，所有需要在构造时传入
     */
    Class<T> clazz;

    public ExcelExportResultVO(Class<T> clazz) {
        this.clazz = clazz;
    }
}
