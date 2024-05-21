package com.antaiib.framework.excel.core.handler.style;

import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.google.common.collect.TreeBasedTable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.Map;

/**
 * 错误数据格式处理器
 *
 * @author yz
 * @since 2024/05/08 10:02
 */
public class ErrorSheetWriteHandler implements RowWriteHandler {

    /**
     * 校验错误文件
     */
    private final TreeBasedTable<Long, Integer, String> errMsgTable;


    public ErrorSheetWriteHandler(TreeBasedTable<Long, Integer, String> errMsgTable) {
        this.errMsgTable = errMsgTable;
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row,
                                Integer relativeRowIndex, Boolean isHead) {
        if (Boolean.TRUE.equals(isHead)) {
            return;
        }

        Sheet sheet = writeSheetHolder.getSheet();
        int rowNum = row.getRowNum();
        // 设置批注
        Map<Integer, String> rowErrMap = errMsgTable.rowMap().get((long) rowNum);
        if (rowErrMap == null) {
            return;
        }
        for (Map.Entry<Integer, String> cellMap : rowErrMap.entrySet()) {
            setComment(sheet, rowNum, cellMap.getKey(), cellMap.getValue());
        }
    }

    /**
     * 设置样式添加批注信息
     */
    private void setComment(Sheet sheet, Integer rowNum, Integer i, String msg) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        // 设置前景填充样式
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 设置前景色为红色
        cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        // 设置垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
        // 创建一个批注
        Comment comment =
                drawingPatriarch.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 1, 1));
        // 输入批注信息
        comment.setString(new XSSFRichTextString(msg));
        // 将批注添加到单元格对象中
        Cell cell = sheet.getRow(rowNum).getCell(i);
        cell.setCellComment(comment);
        cell.setCellStyle(cellStyle);
    }

}
