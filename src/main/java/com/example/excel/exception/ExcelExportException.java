package com.example.excel.exception;

/**
 * Excel导出异常
 * 统一处理Excel导出过程中的所有异常
 * 
 * @author Excel Export Tool
 * @version 1.0.0
 */
public class ExcelExportException extends RuntimeException {

    /**
     * 错误码
     */
    private String errorCode;

    public ExcelExportException(String message) {
        super(message);
    }

    public ExcelExportException(String message, String errorCode) {
        super(message);
        this.errorCode = errorCode;
    }

    public ExcelExportException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelExportException(String message, String errorCode, Throwable cause) {
        super(message, cause);
        this.errorCode = errorCode;
    }

    public String getErrorCode() {
        return errorCode;
    }
}
