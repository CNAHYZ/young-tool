package com.antaiib.framework.excel.core.vo;

import com.antaiib.framework.excel.core.annotations.ExcelLine;
import lombok.Data;
import lombok.experimental.Accessors;

/**
 * excel导入入参VO基类
 *
 * @author yz
 * @since 2024/5/7 9:06
 */

@Data
@Accessors(chain = false)
public class ExcelImportBaseReqVO {

    /**
     * 行号
     */
    @ExcelLine
    public Long oRowNum;
}
