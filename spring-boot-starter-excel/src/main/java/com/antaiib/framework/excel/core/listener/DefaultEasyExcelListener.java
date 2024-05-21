package com.antaiib.framework.excel.core.listener;

import cn.hutool.core.date.DatePattern;
import cn.hutool.core.date.DateUtil;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.util.StringUtils;
import com.antaiib.framework.excel.core.annotations.ExcelLine;
import com.antaiib.framework.excel.core.constants.ErrorExportTypeEnum;
import com.antaiib.framework.excel.core.error.ExcelError;
import com.antaiib.framework.excel.core.util.ExcelUtils;
import com.antaiib.framework.excel.core.validator.ExcelValidateHelper;
import com.antaiib.framework.excel.core.vo.ExcelImportBaseReqVO;
import com.google.common.collect.TreeBasedTable;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicLong;

import static com.antaiib.framework.common.exception.util.ServiceExceptionUtil.exception;

/**
 * 默认easy-excel监听器
 *
 * @author yz
 * @since 2024/05/06 19:51
 */
@Slf4j
public class DefaultEasyExcelListener<T extends ExcelImportBaseReqVO> extends AbstractAnalysisEventListener<T> {
    private Long lineNum = 0L;

    private Integer tempSize = 0;

    /**
     * 额外参数，可传入至checkImportExcel()方法中
     */
    private Map<String, String[]> parameterMap;

    public DefaultEasyExcelListener(Class<T> clazz) {
        super(clazz);
    }

    /**
     * 这个每一条数据解析都会来调用
     *
     * @param t               one row value. Is is same as {@link AnalysisContext#readRowHolder()}
     * @param analysisContext analysisContext
     */
    @Override
    public void invoke(T t, AnalysisContext analysisContext) {
        lineNum++;
        tempSize++;
        Map<Integer, String> resultMap;
        try {
            // 根据excel数据实体中的javax.validation + 正则表达式来校验excel数据
            resultMap = ExcelValidateHelper.validateEntity(t);
        } catch (NoSuchFieldException e) {
            throw new ExcelAnalysisException("第" + analysisContext.readRowHolder().getRowIndex() + "行解析数据出错");
        }
        // 为ExcelLine注解的字段设置行号
        setValueForExcelLineField(t);
        if (resultMap != null) {
            errDTO.add(t, resultMap);
        } else {
            successList.add(t);
        }
        // 默认每1000条处理一次
        if (enableSharding && tempSize >= shardingSize && excelDataChecker != null) {
            excelDataChecker.checkImportExcel(successList, errDTO, parameterMap);
            tempSize = 1;
        }
    }

    /**
     * 所有数据解析完成了 都会来调用
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.debug("Excel read analysed");
        if (excelDataChecker != null) {
            excelDataChecker.checkImportExcel(successList, errDTO, parameterMap);
        }

        try {
            exportErrorExcel();
        } catch (IOException e) {
            log.error("导出错误excel失败", e);
        }
    }


    /**
     * 校验excel头部格式，必须完全匹配
     *
     * @param headMap 传入excel的头部（第一行数据）数据的index,name
     * @param context 上下文
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        if (clazz == null) {
            throw exception(ExcelError.EXCEL_0001, "系统未知异常！");
        }

        try {
            Map<Integer, String> indexNameMap = getIndexNameMap(clazz);
            indexNameMap.forEach((k, v) -> {
                if (StringUtils.isEmpty(headMap.get(k))) {
                    throw exception(ExcelError.EXCEL_0004);
                }
                if (!headMap.get(k).equals(v)) {
                    throw exception(ExcelError.EXCEL_0006);
                }
            });
        } catch (NoSuchFieldException e) {
            log.error("系统异常，解析excel出错，请传入正确格式的excel", e);
        }
    }

    /**
     * 获取注解里ExcelProperty的value，用作校验excel
     *
     * @param clazz clazz
     * @return {@link Map}<{@link Integer}, {@link String}>
     * @throws NoSuchFieldException 无此字段异常
     */
    public Map<Integer, String> getIndexNameMap(Class<T> clazz) throws NoSuchFieldException {
        Map<Integer, String> result = new HashMap<>();
        Field field;
        Field[] fields = clazz.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            field = clazz.getDeclaredField(fields[i].getName());
            field.setAccessible(true);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null) {
                int index = excelProperty.index();
                index = index == -1 ? i : index;
                String[] values = excelProperty.value();
                StringBuilder value = new StringBuilder();
                for (String v : values) {
                    value.append(v);
                }
                result.put(index, value.toString());
            }
        }
        return result;
    }


    /**
     * 错误数据导出
     */
    private void exportErrorExcel() throws IOException {
        if (!ErrorExportTypeEnum.isExportErrorData(errorExportType)) {
            return;
        }
        // 导出结果集
        Collection<T> errDataList;
        // 行-列-错误信息
        TreeBasedTable<Long, Integer, String> errMsgTable;
        if (errorExportType == ErrorExportTypeEnum.ONLY_ERROR_DATA) {
            errMsgTable = TreeBasedTable.create();
            errDataList = errDTO.getErrDataMap().values();

            AtomicLong index = new AtomicLong(1);
            errDTO.getErrMsgTable().rowMap().forEach((k, v) -> {
                Map<Integer, String> row = errMsgTable.row(index.get());
                row.putAll(v);
                index.incrementAndGet();
            });
        } else {
            errDataList = new TreeSet<>(Comparator.comparing(ExcelImportBaseReqVO::getORowNum));
            errDataList.addAll(successList);
            errDataList.addAll(errDTO.getErrDataMap().values());
            errMsgTable = errDTO.getErrMsgTable();
        }
        if (!errMsgTable.isEmpty()) {
            // 导出excel
            ExcelUtils.webWriteErrorExcel(response, errDataList, clazz, errMsgTable,
                    "导入错误信息" + DateUtil.format(new Date(), DatePattern.PURE_DATETIME_MS_PATTERN) + ".xlsx",
                    "导入错误信息");
        }
    }


    /**
     * 为ExcelLine注解的字段设置行号
     *
     * @param t t
     */
    private void setValueForExcelLineField(T t) {
        Field[] fields = t.getClass().getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelLine.class) && field.getType() == Long.class) {
                try {
                    field.setAccessible(true);
                    field.set(t, lineNum);
                } catch (IllegalAccessException e) {
                    log.error("设置ExcelLine注解的字段行号失败", e);
                }
            }
        }
        t.setORowNum(lineNum);
    }
}