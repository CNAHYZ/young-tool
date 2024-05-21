package com.antaiib.framework.excel.core.processor.sheet;


import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.antaiib.framework.excel.core.annotations.ResponseExcel;
import com.antaiib.framework.excel.core.annotations.Sheet;
import com.antaiib.framework.excel.core.config.ExcelConfigProperties;
import com.antaiib.framework.excel.core.enhance.WriterBuilderEnhancer;
import com.antaiib.framework.excel.core.error.ExcelException;
import com.antaiib.framework.excel.core.handler.AbstractSheetWriteHandler;
import com.antaiib.framework.excel.core.vo.ExcelExportResultVO;
import org.springframework.beans.factory.ObjectProvider;

import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;


/**
 * @author pig-mesh
 * @date 2023/09/08
 */
public class ManySheetWriteProcessor extends AbstractSheetWriteHandler {

    public ManySheetWriteProcessor(ExcelConfigProperties configProperties,
                                   ObjectProvider<List<Converter<?>>> converterProvider, WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * 当且仅当List不为空且List中的元素也是List 才返回true
     *
     * @param obj 返回对象
     * @return boolean
     */
    @Override
    public boolean support(Object obj) {
        if (obj instanceof List) {
            List<?> objList = (List<?>) obj;
            return !objList.isEmpty() && objList.get(0) instanceof List;
        } else {
            throw new ExcelException("@ResponseExcel 返回值必须为List<ExcelExportResultVO>类型");
        }
    }

    @Override
    public void write(Object obj, HttpServletResponse response, ResponseExcel responseExcel) {
        if (!(obj instanceof List) || !(((List<?>) obj).get(0) instanceof ExcelExportResultVO)) {
            throw new ExcelException("@ResponseExcel 返回值必须为List<ExcelExportResultVO>类型");
        }
        ExcelWriterBuilder excelWriterBuilder = getExcelWriterBuilder(response, responseExcel);
        // 如果有下拉选，则注册下拉选处理器
        Map<Integer, List<String>> singleSelectorMap = ((ExcelExportResultVO<?>) obj).getSingleSelectorMap();
        if (CollUtil.isNotEmpty(singleSelectorMap)) {
            //     TODO 多sheet暂未处理
        }
        ExcelWriter excelWriter = excelWriterBuilder.build();
        List<?> list = ((ExcelExportResultVO<?>) obj).getData();

        Sheet[] sheets = responseExcel.sheets();
        WriteSheet sheet;
        for (int i = 0; i < sheets.length; i++) {
            List<?> eleList = (List<?>) list.get(i);
            Class<?> dataClass = eleList.get(0).getClass();
            // 创建sheet
            sheet = this.sheet(sheets[i], dataClass, responseExcel.template(), responseExcel.headGenerator());
            // 填充 sheet
            if (responseExcel.fill()) {
                excelWriter.fill(eleList, sheet);
            } else {
                // 写入sheet
                excelWriter.write(eleList, sheet);
            }
        }
        excelWriter.finish();
    }

}
