package com.antaiib.framework.excel.core.handler.style;

import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import java.lang.reflect.Field;

/**
 * 公用单元格写入处理程序
 *
 * @author yz
 * @since 2024/05/07 22:22
 */
@Data
public class CommonCellWriteHandler implements CellWriteHandler {

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        WriteCellData<?> cellData = context.getFirstCellData();
        WriteCellStyle writeCellStyle = cellData.getOrCreateStyle();

        if (Boolean.TRUE.equals(context.getHead())) {
            Cell cell = context.getCell();
            // 设置标题字体样式
            WriteFont headWriteFont = new WriteFont();
            // 保留原来的样式，否则字体样式会丢
            WriteFont.merge(writeCellStyle.getWriteFont(), headWriteFont);

            Field field = context.getHeadData().getField();
            // 必填项标红
            if (field.isAnnotationPresent(NotBlank.class) || field.isAnnotationPresent(NotNull.class)) {
                // 设置字体颜色
                headWriteFont.setColor(IndexedColors.RED.getIndex());

                // 添加必填批注，创建绘图对象
                Drawing<?> drawing = cell.getSheet().createDrawingPatriarch();
                Comment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, cell.getColumnIndex(), 0, cell.getColumnIndex(), 1));
                comment.setString(new XSSFRichTextString("注意：表头红字的列为必填项！"));
                cell.setCellComment(comment);
            }
            writeCellStyle.setWriteFont(headWriteFont);
            // 设置行高
            cell.getRow().setHeightInPoints(40);
        }
    }
}
