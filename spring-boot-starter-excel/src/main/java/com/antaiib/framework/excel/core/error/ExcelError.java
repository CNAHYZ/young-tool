package com.antaiib.framework.excel.core.error;


import com.antaiib.framework.common.exception.ErrorCode;

/**
 * excel错误
 *
 * @author panyc
 * @since 2024/05/09 14:15
 */
public interface ExcelError {
   ErrorCode EXCEL_0001 = new ErrorCode(1, "导入错误，[{}]");
   ErrorCode EXCEL_0002 = new ErrorCode(2, "导入模版错误，导入失败");
   ErrorCode EXCEL_0003 = new ErrorCode(3, "导入数据为空，导入失败");
   ErrorCode EXCEL_0004 = new ErrorCode(4, "解析失败，excel模板或数据有误");
   ErrorCode EXCEL_0005 = new ErrorCode(5, "文件有误，导入失败！");
   ErrorCode EXCEL_0006 = new ErrorCode(6, "解析excel出错，请下载最新模板导入！");

}
