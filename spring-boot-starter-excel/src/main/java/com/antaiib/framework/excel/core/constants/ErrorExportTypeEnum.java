package com.antaiib.framework.excel.core.constants;

import lombok.Getter;

/**
 * 错误数据导出类型枚举
 *
 * @author yz
 * @since 2024/5/7 9:43
 */
@Getter
public enum ErrorExportTypeEnum {
    /**
     * 不自动导出错误数据（由开发者自行处理）
     */
    NONE(0),

    /**
     * 仅自动导出错误的数据
     */
    ONLY_ERROR_DATA(1),

    /**
     * 导出所有数据（包含正确的数据+错误的数据）
     */
    ALL_DATA(2);

    private final Integer code;


    ErrorExportTypeEnum(Integer code) {
        this.code = code;
    }

    public static boolean isExportErrorData(ErrorExportTypeEnum errorExportType) {
        return errorExportType == ONLY_ERROR_DATA || errorExportType == ALL_DATA;
    }

}
