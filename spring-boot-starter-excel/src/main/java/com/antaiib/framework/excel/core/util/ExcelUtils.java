package com.antaiib.framework.excel.core.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.converters.longconverter.LongStringConverter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.antaiib.framework.excel.core.constants.ExcelConst;
import com.antaiib.framework.excel.core.error.ExcelException;
import com.antaiib.framework.excel.core.handler.style.CommonHorizontalCellStyleStrategy;
import com.antaiib.framework.excel.core.handler.style.ErrorSheetWriteHandler;
import com.google.common.collect.TreeBasedTable;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.MediaTypeFactory;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * @author panyc
 * @className ExcelUtils
 * @describe
 * @date 2023/10/11 16:54
 */
@Slf4j
public class ExcelUtils {
    public static void exportExcel(HttpServletResponse response, List<Map<String, Object>> mapList, String fileName) throws Exception {
        Map<String, Object> map = mapList.get(0);
        String[] headArray = new String[map.keySet().size()];
        List<List<Object>> dataList = new ArrayList<>();
        Iterator<String> iterator = map.keySet().iterator();
        int i = 0;
        while (iterator.hasNext()) {
            headArray[i] = iterator.next();
            i++;
        }
        for (Map<String, Object> map1 : mapList) {
            List<Object> data = new ArrayList<>();
            for (int j = 0; j < headArray.length; j++) {
                data.add(map1.get(headArray[j]));
            }
            dataList.add(data);
        }
        setResponseHeaders(response, fileName);
        EasyExcel.write(response.getOutputStream()).head(createdHead(headArray)).sheet(fileName).doWrite(dataList);
    }

    public static <T> void write(
            HttpServletResponse response, String filename, String sheetName,
            Class<T> head, List<T> data,
            WriteHandler... writeHandlers) { // 可变参数用于不确定数量的写处理器
        try {
            // 设置响应头信息
            setResponseHeaders(response, filename);

            // 创建EasyExcel写操作的构建器
            ExcelWriterBuilder writerBuilder = EasyExcel.write(response.getOutputStream(), head);
            // 注册写处理器
            for (WriteHandler writeHandler : writeHandlers) {
                writerBuilder.registerWriteHandler(writeHandler);
            }

            // 执行写操作
            // 动态传入的sheet名称
            writerBuilder.sheet(sheetName)
                    // 动态传入的数据列表
                    .doWrite(data);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void setResponseHeaders(HttpServletResponse response, String filename) throws UnsupportedEncodingException {
        // 根据实际的文件类型找到对应的 contentType
        String contentType = MediaTypeFactory.getMediaType(filename).map(MediaType::toString)
                .orElse("application/vnd.ms-excel");
        response.setContentType(contentType);
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=" + URLEncoder.encode(filename, StandardCharsets.UTF_8.name()));
    }

    /**
     * 将列表以 Excel 响应给前端
     *
     * @param response  响应
     * @param filename  文件名
     * @param sheetName Excel sheet 名
     * @param head      Excel head 头
     * @param data      数据列表哦
     * @param <T>       泛型，保证 head 和 data 类型的一致性
     * @throws IOException 写入失败的情况
     */
    public static <T> void write(HttpServletResponse response, String filename, String sheetName,
                                 Class<T> head, List<T> data) throws IOException {
        // 输出 Excel
        EasyExcel.write(response.getOutputStream(), head)
                // 不要自动关闭，交给 Servlet 自己处理
                .autoCloseStream(false)
                // 基于 column 长度，自动适配。最大 255 宽度
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                // 避免 Long 类型丢失精度
                .registerConverter(new LongStringConverter())
                .sheet(sheetName).doWrite(data);
        // 设置 header 和 contentType。写在最后的原因是，避免报错时，响应 contentType 已经被修改了
        setResponseHeaders(response, filename);
    }

    /**
     * web写入错误信息的excel
     *
     * @throws IOException IOException
     */
    public static void webWriteErrorExcel(HttpServletResponse response, Collection<?> objects, Class<?> clazz,
                                          TreeBasedTable<Long, Integer, String> errMsgTable, String fileName, String sheetName) throws IOException {
        setResponseHeaders(response, fileName);

        try (ServletOutputStream outputStream = response.getOutputStream()) {
            EasyExcel.write(outputStream, clazz)
                    .inMemory(Boolean.TRUE)
                    // 注册默认行样式
                    .registerWriteHandler(CommonHorizontalCellStyleStrategy.STRATEGY)
                    // 注册默认列宽样式
                    .registerWriteHandler(new SimpleColumnWidthStyleStrategy(ExcelConst.HEAD_COLUMN_WIDTH))
                    .registerWriteHandler(new ErrorSheetWriteHandler(errMsgTable))
                    .sheet(sheetName)
                    .doWrite(objects);
        } catch (Exception e) {
            log.error("webWriteErrorExcel error", e);
        }
    }

    public static <T> List<T> read(MultipartFile file, Class<T> head) throws IOException {
        return EasyExcel.read(file.getInputStream(), head, null)
                // 不要自动关闭，交给 Servlet 自己处理
                .autoCloseStream(false)
                .doReadAllSync();
    }

    private static List<List<String>> createdHead(String[] headMap) {
        List<List<String>> headList = new ArrayList<List<String>>();
        for (String head : headMap) {
            List<String> list = new ArrayList<String>();
            list.add(head);
            headList.add(list);
        }
        return headList;
    }

    /**
     * @param offset   偏移量，如果给0，表示从A列开始，1，就是从B列
     * @param rowId    第几行
     * @param colCount 一共多少列
     * @return 如果给入参 1,1,10. 表示从B1-K1。最终返回 $B$1:$K$1
     * @author denggonghai 2016年8月31日 下午5:17:49
     */
    public static String getRange(int offset, int rowId, int colCount) {
        char start = (char) ('A' + offset);
        if (colCount <= 25) {
            char end = (char) (start + colCount - 1);
            return "$" + start + "$" + rowId + ":$" + end + "$" + rowId;
        } else {
            char endPrefix = 'A';
            char endSuffix = 'A';
            // 26-51之间，包括边界（仅两次字母表计算）
            if ((colCount - 25) / 26 == 0 || colCount == 51) {
                // 边界值
                if ((colCount - 25) % 26 == 0) {
                    endSuffix = (char) ('A' + 25);
                } else {
                    endSuffix = (char) ('A' + (colCount - 25) % 26 - 1);
                }
            } else {// 51以上
                if ((colCount - 25) % 26 == 0) {
                    endSuffix = (char) ('A' + 25);
                    endPrefix = (char) (endPrefix + (colCount - 25) / 26 - 1);
                } else {
                    endSuffix = (char) ('A' + (colCount - 25) % 26 - 1);
                    endPrefix = (char) (endPrefix + (colCount - 25) / 26);
                }
            }
            return "$" + start + "$" + rowId + ":$" + endPrefix + endSuffix + "$" + rowId;
        }
    }


    /**
     * 设置验证规则
     *
     * @param sheet       sheet对象
     * @param helper      验证助手
     * @param constraint  createExplicitListConstraint
     * @param addressList 验证位置对象
     * @param msgHead     错误提示头
     * @param msgContext  错误提示内容
     */
    private static void setValidation(Sheet sheet, DataValidationHelper helper, DataValidationConstraint constraint, CellRangeAddressList addressList, String msgHead, String msgContext) {
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        dataValidation.setShowErrorBox(true);
        dataValidation.setSuppressDropDownArrow(true);
        dataValidation.createErrorBox(msgHead, msgContext);
        sheet.addValidationData(dataValidation);
    }


    /**
     * 设置下拉列表的长度，excel有长度限制，255字节,需要新建立一个sheet页存储数据，随后隐藏
     *
     * @param sheet
     * @param sheetName
     * @param firstRow  起始行
     * @param endRow    终止行
     * @param firstCol  起始列
     * @param endCol    终止列
     * @return
     */
    public static DataValidation setDataValidation(Sheet sheet, String sheetName, int firstRow, int endRow, int firstCol, int endCol, String[] arr) {
        Workbook workbook = sheet.getWorkbook();
        // 建立新的sheet页存储指定列的下拉列表数据
        Sheet hiddenSheet = workbook.createSheet(sheetName);
        for (int i = 0; i < arr.length; i++) {
            Row row = hiddenSheet.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(arr[i]);
        }
        String strFormula = sheetName + "!$A$1:$A$" + arr.length;
        if (arr.length < 1) {
            strFormula = sheetName + "!$A$1:$A$" + 1;
        }
        // 原顺序为 起始行 起始列 终止行 终止列
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint formulaListConstraint = dvHelper.createFormulaListConstraint(strFormula);
        return dvHelper.createValidation(formulaListConstraint, regions);
    }


    /**
     * 获取字段所在的列号
     *
     * @param clazz clazz
     * @return {@link Map}<{@link Integer}, {@link String}>
     */
    public static Map<String, Integer> getFieldIndexMap(Class<?> clazz) {
        Map<String, Integer> result = new HashMap<>();
        Field field;
        Field[] fields = clazz.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            try {
                field = clazz.getDeclaredField(fields[i].getName());
            } catch (NoSuchFieldException e) {
                throw new ExcelException(e);
            }
            field.setAccessible(true);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null) {
                int index = excelProperty.index();
                index = index == -1 ? i : index;
                result.put(field.getName(), index);
            }
        }
        return result;
    }
}
