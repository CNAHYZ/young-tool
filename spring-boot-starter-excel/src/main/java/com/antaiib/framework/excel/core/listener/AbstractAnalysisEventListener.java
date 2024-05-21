package com.antaiib.framework.excel.core.listener;

import com.alibaba.excel.event.AnalysisEventListener;
import com.antaiib.framework.excel.core.constants.ErrorExportTypeEnum;
import com.antaiib.framework.excel.core.validator.IExcelDataChecker;
import com.antaiib.framework.excel.core.vo.ExcelImportBaseReqVO;
import com.antaiib.framework.excel.core.vo.ExcelImportErrDto;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;


/**
 * list analysis EventListener
 *
 * @author pig-mesh
 * @date 2023/09/08
 */
@NoArgsConstructor
public abstract class AbstractAnalysisEventListener<T extends ExcelImportBaseReqVO> extends AnalysisEventListener<T> {
    /**
     * 错误数据结果集
     */
    @Getter
    ExcelImportErrDto<T> errDTO = new ExcelImportErrDto<>();

    /**
     * 成功数据结果集
     */
    @Getter
    final List<T> successList = new ArrayList<>();

    /**
     * 存放解析的临时对象
     */
    final List<T> dataList = new ArrayList<>();

    /**
     * excel内容检查器
     */
    @Setter
    IExcelDataChecker<T> excelDataChecker;

    /**
     * 每次分片导入的大小，默认每隔1000条存储数据库，然后清理list ，方便内存回收
     */
    @Setter
    int shardingSize = 1000;

    /**
     * 是否开启分片导入
     */
    @Setter
    boolean enableSharding = false;

    /**
     * 错误数据导出类型
     */
    @Setter
    ErrorExportTypeEnum errorExportType = ErrorExportTypeEnum.ONLY_ERROR_DATA;

    /**
     * response
     */
    @Setter
    HttpServletResponse response;

    /**
     * excel对应的实体对象的反射类
     */
    @Setter
    Class<T> clazz;

    AbstractAnalysisEventListener(int shardingSize, ErrorExportTypeEnum errorExportType, IExcelDataChecker<T> excelDataChecker,
                                  HttpServletResponse response) {
        this.shardingSize = shardingSize;
        this.errorExportType = errorExportType;
        this.excelDataChecker = excelDataChecker;
        this.response = response;
    }

    AbstractAnalysisEventListener(Class<T> clazz) {
        this.clazz = clazz;
    }

}
