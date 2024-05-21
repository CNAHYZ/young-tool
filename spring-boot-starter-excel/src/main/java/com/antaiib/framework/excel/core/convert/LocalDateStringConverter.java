
package com.antaiib.framework.excel.core.convert;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;



/**
 * 时间转换
 * @author  pig-mesh
 * @date 2023/09/08
 */

public enum LocalDateStringConverter implements Converter<LocalDate> {
    /**
     * 实例
     */

    INSTANCE;

    @Override
    public Class<?> supportJavaTypeKey() {
        return LocalDate.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public LocalDate convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        if (contentProperty == null || contentProperty.getDateTimeFormatProperty() == null) {
            return LocalDate.parse(cellData.getStringValue());
        } else {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(
                    contentProperty.getDateTimeFormatProperty().getFormat());
            return LocalDate.parse(cellData.getStringValue(), formatter);
        }
    }

    @Override
    public WriteCellData<?> convertToExcelData(LocalDate value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        DateTimeFormatter formatter;
        if (contentProperty == null || contentProperty.getDateTimeFormatProperty() == null) {
            formatter = DateTimeFormatter.ISO_LOCAL_DATE;
        } else {
            formatter = DateTimeFormatter.ofPattern(contentProperty.getDateTimeFormatProperty().getFormat());
        }
        return new WriteCellData<>(value.format(formatter));
    }
}

