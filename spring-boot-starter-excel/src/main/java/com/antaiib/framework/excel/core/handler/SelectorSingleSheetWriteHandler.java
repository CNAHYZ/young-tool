package com.antaiib.framework.excel.core.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.List;
import java.util.Map;

/**
 * 下拉框-无联动-单sheet，多sheet暂未支持
 *
 * @author panyc
 * @since 2024/05/07 17:04
 */
@AllArgsConstructor
public class SelectorSingleSheetWriteHandler extends AbstractSelectorSheetWriteHandler {
    /**
     * 列的索引-数据下拉选项映射
     */
    private Map<Integer, List<String>> indexDataMap;

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(indexDataMap)) {
            return;
        }

        DataValidationHelper helper = writeSheetHolder.getSheet().getDataValidationHelper();

        Workbook workbook = writeWorkbookHolder.getWorkbook();
        // 循环创建下拉选
        indexDataMap.forEach((columnIndex, dateList) -> {
            // 创建一个隐藏的sheet
            String sheetName = "SHEET_MAP2" + columnIndex;
            Sheet tempSheet = workbook.createSheet(sheetName);
            // i:表示你开始的行数0表示你开始的列数
            for (int i = 0, length = dateList.size(); i < length; i++) {
                tempSheet.createRow(i).createCell(0).setCellValue(dateList.get(i));
            }
            Name categoryName = workbook.createName();
            categoryName.setNameName(sheetName);
            //$A$1:$A$N代表以A列1行开始获取N行下拉数据
            categoryName.setRefersToFormula(sheetName + "!$A$1:$A$" + (dateList.size()));
            // 5将刚才设置的sheet引用到你的下拉列表中
            CellRangeAddressList addressList = new CellRangeAddressList(1, 999, columnIndex, columnIndex);
            DataValidationConstraint constraint = helper.createFormulaListConstraint(sheetName);

            setValidation(writeSheetHolder.getSheet(), helper, constraint, addressList, "警告", "请下拉选择合适的值！");

            // 将新建立的sheet页隐藏掉
            int sheetIndex = workbook.getSheetIndex(tempSheet);
            writeWorkbookHolder.getWorkbook().setSheetHidden(sheetIndex, true);
        });

    }
}

