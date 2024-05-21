package com.antaiib.framework.excel.core.vo;

import com.google.common.collect.TreeBasedTable;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Map;
import java.util.TreeMap;

/**
 * excel数据导入错误封装对象
 *
 * @author yz
 * @since 2024/05/08 09:11
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelImportErrDto<T extends ExcelImportBaseReqVO> {
    /**
     * 行号-原始错误数据
     */
    protected Map<Long, T> errDataMap = new TreeMap<>(Long::compareTo);
    /**
     * 行号-列号-错误信息
     */
    protected TreeBasedTable<Long, Integer, String> errMsgTable = TreeBasedTable.create();

    public ExcelImportErrDto<T> add(T errData, int colNum, String errMsg) {
        errDataMap.put(errData.getORowNum(), errData);
        errMsgTable.put(errData.getORowNum(), colNum, errMsg);
        return this;
    }

    /**
     * 添加
     *
     * @param errData   错误数据
     * @param errMsgMap 列号-错误消息的映射
     * @return {@link ExcelImportErrDto}<{@link T}>
     */
    public ExcelImportErrDto<T> add(T errData, Map<Integer, String> errMsgMap) {
        errDataMap.put(errData.getORowNum(), errData);
        errMsgMap.forEach((k, v) -> errMsgTable.put(errData.getORowNum(), k, v));
        return this;
    }

    public ExcelImportErrDto<T> add(ExcelImportErrDto<T> errDto) {
        errDataMap.putAll(errDto.getErrDataMap());
        errMsgTable.putAll(errDto.getErrMsgTable());
        return this;
    }
}
