package com.antaiib.framework.excel.core.error;


/**
 * @author pig-mesh
 * @date 2023/09/08
 */
public class ExcelException extends RuntimeException {
    
    private static final long serialVersionUID = 1L;
    
    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(NoSuchFieldException e) {
        super(e);
    }
}
