package com.antaiib.framework.excel.core.handler;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.antaiib.framework.excel.core.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author panyc
 * @className AbstractSelectorSheetWriteHandler
 * @describe
 * @date 2023/10/13 15:43
 */
public class AbstractSelectorSheetWriteHandler implements SheetWriteHandler {
    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    /**
     * 设置验证规则
     * @param sheet			sheet对象
     * @param helper		验证助手
     * @param constraint	createExplicitListConstraint
     * @param addressList	验证位置对象
     * @param msgHead		错误提示头
     * @param msgContext	错误提示内容
     */
    protected void setValidation(Sheet sheet, DataValidationHelper helper, DataValidationConstraint constraint, CellRangeAddressList addressList, String msgHead, String msgContext) {
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        dataValidation.setShowErrorBox(true);
        dataValidation.setSuppressDropDownArrow(true);
        dataValidation.createErrorBox(msgHead, msgContext);
        sheet.addValidationData(dataValidation);
    }

    protected static void writeData(Workbook hssfWorkBook, Sheet mapSheet, List<String> provinceList, Map<String, List<String>> siteMap) {
        //循环将父数据写入siteSheet的第1行中
        int siteRowId = 0;
        Row provinceRow = mapSheet.createRow(siteRowId++);
        provinceRow.createCell(0).setCellValue("父列表");
        for (int i = 0; i < provinceList.size(); i++) {
            //有多少个省，创建多少个下拉框
            provinceRow.createCell(i + 1).setCellValue(provinceList.get(i));
        }
        // 将具体的数据写入到每一行中，行开头为父级区域，后面是子区域。
        Iterator<String> keyIterator = siteMap.keySet().iterator();
        while (keyIterator.hasNext()) {
            String key = keyIterator.next();
            List<String> son = siteMap.get(key);
            Row siteRow = mapSheet.createRow(siteRowId++);

            siteRow.createCell(0).setCellValue(key);
            for (int i = 0; i < son.size(); i++) {
                siteRow.createCell(i + 1).setCellValue(son.get(i));

            }

            // 添加名称管理器
            String range = ExcelUtils.getRange(1, siteRowId, son.size());
            Name name = hssfWorkBook.createName();
            name.setNameName(key);
            String formula = mapSheet.getSheetName() + "!" + range;
            name.setRefersToFormula(formula);
        }
    }
}
