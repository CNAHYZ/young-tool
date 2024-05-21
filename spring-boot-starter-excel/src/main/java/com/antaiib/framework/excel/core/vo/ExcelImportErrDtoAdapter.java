package com.antaiib.framework.excel.core.vo;

import cn.hutool.core.lang.func.Func1;
import cn.hutool.core.lang.func.LambdaUtil;
import com.antaiib.framework.excel.core.util.ExcelUtils;
import lombok.AccessLevel;
import lombok.EqualsAndHashCode;
import lombok.Getter;

import java.util.Map;

/**
 * excel数据导入错误封装对象
 *
 * @author yz
 * @since 2024/05/08 09:11
 */
@EqualsAndHashCode(callSuper = true)
public class ExcelImportErrDtoAdapter<T extends ExcelImportBaseReqVO> extends ExcelImportErrDto<T> {
    @Getter(AccessLevel.PRIVATE)
    Map<String, Integer> fieldIndexMap;

    public ExcelImportErrDtoAdapter(ExcelImportErrDto<T> excelImportErrDto, Class<T> clazz) {
        super(excelImportErrDto.getErrDataMap(), excelImportErrDto.getErrMsgTable());
        this.fieldIndexMap = ExcelUtils.getFieldIndexMap(clazz);
    }

    public ExcelImportErrDtoAdapter<T> add(T errData, String fieldName, String errMsg) {
        int colNum = fieldIndexMap.get(fieldName);
        errDataMap.put(errData.getORowNum(), errData);
        errMsgTable.put(errData.getORowNum(), colNum, errMsg);
        return this;
    }

    public ExcelImportErrDtoAdapter<T> add(T errData, Func1<T, ?> func, String errMsg) {
        int colNum = fieldIndexMap.get(LambdaUtil.getFieldName(func));
        errDataMap.put(errData.getORowNum(), errData);
        errMsgTable.put(errData.getORowNum(), colNum, errMsg);
        return this;
    }

}