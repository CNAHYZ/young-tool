package com.antaiib.framework.excel.core.vo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;

/**
 * excel数据导入结果VO类
 *
 * @author Administrator
 * @since 2024/05/08 09:11
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExcelImportResultVO<T extends ExcelImportBaseReqVO> {

    /**
     * 成功数据结果集
     */
    private List<T> successList;

    /**
     * 错误数据结果集
     */
    private ExcelImportErrDto<T> errDto;

    public ExcelImportResultVO(ExcelImportErrDto<T> errDto){
        this.successList =new ArrayList<>();
        this.errDto = errDto;
    }
}
