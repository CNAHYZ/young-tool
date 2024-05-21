
package com.antaiib.framework.excel.core.aop;


import cn.hutool.core.util.ReflectUtil;
import cn.hutool.extra.spring.SpringUtil;
import com.alibaba.excel.EasyExcel;
import com.antaiib.framework.common.exception.ServiceException;
import com.antaiib.framework.excel.core.annotations.RequestExcel;
import com.antaiib.framework.excel.core.convert.LocalDateStringConverter;
import com.antaiib.framework.excel.core.convert.LocalDateTimeStringConverter;
import com.antaiib.framework.excel.core.error.ExcelError;
import com.antaiib.framework.excel.core.listener.AbstractAnalysisEventListener;
import com.antaiib.framework.excel.core.listener.DefaultEasyExcelListener;
import com.antaiib.framework.excel.core.validator.IExcelDataChecker;
import com.antaiib.framework.excel.core.vo.ExcelImportResultVO;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.MethodParameter;
import org.springframework.core.ResolvableType;
import org.springframework.web.bind.support.WebDataBinderFactory;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.ModelAndViewContainer;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartRequest;

import javax.annotation.Nonnull;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.lang.reflect.Constructor;

import static com.antaiib.framework.common.exception.util.ServiceExceptionUtil.exception;


/**
 * 上传excel 解析注解
 *
 * @author pig-mesh
 * @date 2023/09/08
 */

@Slf4j
public class RequestExcelArgumentResolver implements HandlerMethodArgumentResolver {

    @Override
    public boolean supportsParameter(MethodParameter parameter) {
        return parameter.hasParameterAnnotation(RequestExcel.class);
    }

    @Override
    @SneakyThrows(Exception.class)
    public Object resolveArgument(MethodParameter parameter, ModelAndViewContainer modelAndViewContainer, @Nonnull NativeWebRequest webRequest, WebDataBinderFactory webDataBinderFactory) {
        Class<?> parameterType = parameter.getParameterType();
        if (!parameterType.isAssignableFrom(ExcelImportResultVO.class)) {
            throw new IllegalArgumentException(
                    "Excel upload request resolver error, @RequestExcel parameter is not ExcelImportResult class: " + parameterType);
        }

        // 处理自定义 readListener
        RequestExcel requestExcel = parameter.getParameterAnnotation(RequestExcel.class);
        assert requestExcel != null;
        Class<? extends IExcelDataChecker> excelChecker = requestExcel.dataChecker();
        Class<? extends AbstractAnalysisEventListener> readListenerClass = requestExcel.readListener();
        // 获取 DefaultEasyExcelListener 类的构造函数
        Constructor<? extends AbstractAnalysisEventListener> constructor = readListenerClass.getConstructor(Class.class);
        Class<?> actualGenericClazz = ResolvableType.forMethodParameter(parameter).getGeneric(0).resolve();
        // 使用构造函数实例化 DefaultEasyExcelListener 类的对象，并传递泛型类型参数的 Class 对象
        AbstractAnalysisEventListener readListener = constructor.newInstance(actualGenericClazz);

        if (readListener instanceof DefaultEasyExcelListener) {
            ReflectUtil.setFieldValue(readListener, "excelDataChecker", SpringUtil.getBean(excelChecker));
            ReflectUtil.setFieldValue(readListener, "response", webRequest.getNativeResponse(HttpServletResponse.class));
            ReflectUtil.setFieldValue(readListener, "clazz", actualGenericClazz);
            ReflectUtil.setFieldValue(readListener, "errorExportType", requestExcel.errorExportType());
            ReflectUtil.setFieldValue(readListener, "shardingSize", requestExcel.shardingSize());
            ReflectUtil.setFieldValue(readListener, "enableSharding", requestExcel.enableSharding());
            ReflectUtil.setFieldValue(readListener, "parameterMap", webRequest.getParameterMap());
        }

        // 获取请求文件流
        HttpServletRequest request = webRequest.getNativeRequest(HttpServletRequest.class);
        assert request != null;
        InputStream inputStream;
        if (request instanceof MultipartRequest) {
            MultipartFile file = ((MultipartRequest) request).getFile(requestExcel.fileName());
            if (file == null) {
                throw exception(ExcelError.EXCEL_0005);
            }
            inputStream = file.getInputStream();
        } else {
            inputStream = request.getInputStream();
        }

        // 获取目标类型
        Class<?> excelModelClass = actualGenericClazz;

        try {
            // 这里需要指定读用哪个 class 去读，然后读取第一个 sheet 文件流会自动关闭
            EasyExcel.read(inputStream, excelModelClass, readListener)
                    .registerConverter(LocalDateStringConverter.INSTANCE)
                    .registerConverter(LocalDateTimeStringConverter.INSTANCE)
                    .ignoreEmptyRow(requestExcel.ignoreEmptyRow())
                    .sheet()
                    .doRead();
        } catch (ServiceException e) {
            throw e;
        } catch (Exception e) {
            Throwable cause = e.getCause();
            if (cause instanceof ServiceException) {
                throw (ServiceException) cause;
            } else {
                log.error("导入失败：", e);
                throw exception(ExcelError.EXCEL_0004);
            }
        }
        return new ExcelImportResultVO<>(readListener.getSuccessList(), readListener.getErrDTO());
    }

}
