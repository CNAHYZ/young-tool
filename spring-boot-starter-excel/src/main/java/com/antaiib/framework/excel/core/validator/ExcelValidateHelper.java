package com.antaiib.framework.excel.core.validator;

import com.alibaba.excel.annotation.ExcelProperty;
import com.antaiib.framework.excel.core.util.ExcelUtils;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.groups.Default;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * excel校验
 *
 * @author yz
 * @since 2024/05/08 19:27
 */
public class ExcelValidateHelper {

    private static final Validator VALIDATOR = Validation.buildDefaultValidatorFactory().getValidator();

    public static <T> Map<Integer, String> validateEntity(T obj) throws NoSuchFieldException, SecurityException {
        Map<Integer, String> resultMap = new HashMap<>();
        Set<ConstraintViolation<T>> set = VALIDATOR.validate(obj, Default.class);
        if (set != null && !set.isEmpty()) {
            Map<String, Integer> fieldIndexMap = ExcelUtils.getFieldIndexMap(obj.getClass());
            for (ConstraintViolation<T> cv : set) {
                Field declaredField = obj.getClass().getDeclaredField(cv.getPropertyPath().toString());
                ExcelProperty annotation = declaredField.getAnnotation(ExcelProperty.class);
                // 拼接错误信息，包含当前出错数据的标题名字+错误信息
                int index = annotation.index();
                if (index == -1) {
                    index = fieldIndexMap.get(cv.getPropertyPath().toString());
                }
                resultMap.put(index, cv.getMessage());
            }
        }

        return resultMap.isEmpty() ? null : resultMap;
    }
}
