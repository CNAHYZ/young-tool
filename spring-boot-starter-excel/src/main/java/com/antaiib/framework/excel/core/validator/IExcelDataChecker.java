package com.antaiib.framework.excel.core.validator;

import com.antaiib.framework.excel.core.vo.ExcelImportBaseReqVO;
import com.antaiib.framework.excel.core.vo.ExcelImportErrDto;

import java.util.List;
import java.util.Map;

/**
 * excel数据检查器
 *
 * @author yz
 * @since 2024/05/08 15:27
 */
public interface IExcelDataChecker<T extends ExcelImportBaseReqVO> {

    /**
     * 校验方法（如有必要（如分片导入），可在该方法直接批量保存到数据库）
     *
     * @param successList  校验成功的数据
     * @param errDtoPara   导入错误封装对象
     * @param parameterMap 接口的其他参数也会传到这里来
     */
    void checkImportExcel(List<T> successList, ExcelImportErrDto<T> errDtoPara, Map<String, String[]> parameterMap);
}
