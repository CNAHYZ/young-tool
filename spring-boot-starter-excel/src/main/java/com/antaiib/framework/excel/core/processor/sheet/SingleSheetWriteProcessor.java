package com.antaiib.framework.excel.core.processor.sheet;


import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.antaiib.framework.excel.core.annotations.ResponseExcel;
import com.antaiib.framework.excel.core.config.ExcelConfigProperties;
import com.antaiib.framework.excel.core.enhance.WriterBuilderEnhancer;
import com.antaiib.framework.excel.core.error.ExcelException;
import com.antaiib.framework.excel.core.handler.AbstractSheetWriteHandler;
import com.antaiib.framework.excel.core.handler.SelectorSingleSheetWriteHandler;
import com.antaiib.framework.excel.core.vo.ExcelExportResultVO;
import org.springframework.beans.factory.ObjectProvider;

import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;


/**
 * 处理单sheet 页面
 *
 * @author pig-mesh
 * @date 2023/09/08
 */
public class SingleSheetWriteProcessor extends AbstractSheetWriteHandler {

    public SingleSheetWriteProcessor(ExcelConfigProperties configProperties,
                                     ObjectProvider<List<Converter<?>>> converterProvider, WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * obj 是List 且list不为空同时list中的元素不是是List 才返回true
     *
     * @param obj 返回对象
     * @return boolean
     */
    @Override
    public boolean support(Object obj) {
        if (obj instanceof ExcelExportResultVO) {
            return true;
        } else {
            throw new ExcelException("@ResponseExcel 返回值必须为ExcelExportResultVO类型");
        }
    }

    @Override
    public void write(Object obj, HttpServletResponse response, ResponseExcel responseExcel) {
        if (!(obj instanceof ExcelExportResultVO)) {
            throw new ExcelException("@ResponseExcel 返回值必须为ExcelExportResultVO类型");
        }
        ExcelWriterBuilder excelWriterBuilder = getExcelWriterBuilder(response, responseExcel);
        // 如果有下拉选，则注册下拉选处理器
        Map<Integer, List<String>> singleSelectorMap = ((ExcelExportResultVO<?>) obj).getSingleSelectorMap();
        if (CollUtil.isNotEmpty(singleSelectorMap)) {
            excelWriterBuilder.registerWriteHandler(new SelectorSingleSheetWriteHandler(singleSelectorMap));
        }
        ExcelWriter excelWriter = excelWriterBuilder.build();
        List<?> list = ((ExcelExportResultVO<?>) obj).getData();
        // 有模板则不指定sheet名
        Class<?> dataClass = ((ExcelExportResultVO<?>) obj).getClazz();
        WriteSheet sheet = this.sheet(responseExcel.sheets()[0], dataClass, responseExcel.template(),
                responseExcel.headGenerator());

        // 填充 sheet
        if (responseExcel.fill()) {
            excelWriter.fill(list, sheet);
        } else {
            // 写入sheet
            excelWriter.write(list, sheet);
        }
        excelWriter.finish();
    }

}
