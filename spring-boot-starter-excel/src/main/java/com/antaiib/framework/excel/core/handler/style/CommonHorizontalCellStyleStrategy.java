package com.antaiib.framework.excel.core.handler.style;

import com.alibaba.excel.metadata.data.DataFormatData;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.antaiib.framework.excel.core.constants.ExcelConst;
import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;


/**
 * 通用横向单元格样式策略
 *
 * @author yz
 * @since 2024/04/29 19:20
 */
@Slf4j
@UtilityClass
public class CommonHorizontalCellStyleStrategy {

    public static final HorizontalCellStyleStrategy STRATEGY = getHorizontalCellStyleStrategy();

    /**
     * 设置excel 单元格样式
     *
     * @return {@link CommonHorizontalCellStyleStrategy}
     * @author zhangkh
     */
    private static HorizontalCellStyleStrategy getHorizontalCellStyleStrategy() {
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 设置表头居中对齐、垂直对其
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        headWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 颜色
        headWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        headWriteCellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());

        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints(ExcelConst.HEAD_CELL_FONT_SIZE);
        // 加粗
        headWriteFont.setBold(true);
        // 字体
        headWriteCellStyle.setWriteFont(headWriteFont);
        // 设置 自动换行
        headWriteCellStyle.setWrapped(true);

        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 设置边框样式
        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
        // 边框颜色
        contentWriteCellStyle.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.index);
        contentWriteCellStyle.setRightBorderColor(IndexedColors.GREY_40_PERCENT.index);
        contentWriteCellStyle.setTopBorderColor(IndexedColors.GREY_40_PERCENT.index);
        contentWriteCellStyle.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.index);
        // 设置内容靠中对齐
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 设置 自动换行
        contentWriteCellStyle.setWrapped(true);
        // 单元格设置为文本格式
        DataFormatData dataFormatData = new DataFormatData();
        dataFormatData.setIndex((short) 49);
        contentWriteCellStyle.setDataFormatData(dataFormatData);
        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        return new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
    }


}
