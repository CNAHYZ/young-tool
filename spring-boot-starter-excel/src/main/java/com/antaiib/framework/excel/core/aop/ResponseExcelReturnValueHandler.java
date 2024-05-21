package com.antaiib.framework.excel.core.aop;

import cn.hutool.core.lang.Assert;
import com.antaiib.framework.excel.core.annotations.ResponseExcel;
import com.antaiib.framework.excel.core.processor.sheet.SheetWriteProcessor;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.MethodParameter;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;

import javax.servlet.http.HttpServletResponse;
import java.util.List;


/**
 * 处理@ResponseExcel 返回值
 * @author pig-mesh
 * @date 2023/09/08
 */
@Slf4j
@RequiredArgsConstructor
public class ResponseExcelReturnValueHandler implements HandlerMethodReturnValueHandler {
    
    private final List<SheetWriteProcessor> sheetWriteProcessorList;
    
    /**
     * 只处理@ResponseExcel 声明的方法
     *
     * @param parameter 方法签名
     * @return 是否处理
     */
    @Override
    public boolean supportsReturnType(MethodParameter parameter) {
        return parameter.getMethodAnnotation(ResponseExcel.class) != null;
    }
    
    /**
     * 处理逻辑
     *
     * @param o                返回参数
     * @param parameter        方法签名
     * @param mavContainer     上下文容器
     * @param nativeWebRequest 上下文
     */
    @Override
    public void handleReturnValue(Object o, MethodParameter parameter, ModelAndViewContainer mavContainer,
            NativeWebRequest nativeWebRequest) {
        /* check */
        HttpServletResponse response = nativeWebRequest.getNativeResponse(HttpServletResponse.class);
        Assert.state(response != null, "No HttpServletResponse");
        ResponseExcel responseExcel = parameter.getMethodAnnotation(ResponseExcel.class);
        Assert.state(responseExcel != null, "No @ResponseExcel");
        mavContainer.setRequestHandled(true);
        
        sheetWriteProcessorList.stream().filter(processor -> processor.support(o)).findFirst()
                .ifPresent(processor -> processor.export(o, response, responseExcel));
    }
    
}
