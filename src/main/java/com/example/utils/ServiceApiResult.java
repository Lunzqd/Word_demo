package com.example.utils;

import lombok.Data;

import java.io.Serializable;

/**
 * @author al
 * @date 2019/5/10 11:24
 * @description service接口数据返回统一标准
 */
@Data
public class ServiceApiResult<T> implements Serializable {

    private static final long serialVersionUID = 1L;

    // 执行成功标志:默认成功
    private Boolean success = true;

    // 执行结果：建议使用 entity | map | list | string
    private T data;

    // 信息
    private String message;

    // 信息
    private String errorMessage;

    // 数据总数
    private int count;

    // Code
    private int code;

    public ServiceApiResult() {

    }

    ServiceApiResult(int code, boolean success, String message, int count, T data) {
        this.code = code;
        this.success = success;
        this.message = message;
        this.data = data;
        this.count = count;
    }

    /**
     * 分页单个数据
     */
    public static <T> ServiceApiResult<T> ok(T data) {
        try {
            ServiceApiResult<T> result = new ServiceApiResult<T>(200, true, "success", 1, data);

            return result;
        } catch (Exception ex) {
            return error(ex);
        }
    }

    public static <T> ServiceApiResult<T> error(Exception ex) {
        String exceptionMessageFormat = "Message: %s, StackTrace: %s, Suppressed: %s, Cause: %s, Class: %s %s";

        String msg = String.format(exceptionMessageFormat, ex.getMessage(), ex.getStackTrace(), ex.getSuppressed(),
                ex.getCause(), ex.getClass(), System.getProperty("line.separator"));

        return error(msg);
    }

    public static <T> ServiceApiResult<T> error(String message) {
        ServiceApiResult<T> result = new ServiceApiResult<T>(500, false, message, 0, null);
        result.count = 0;

        return result;
    }

    public static <T> ServiceApiResult<T> error(int code, String message) {
        ServiceApiResult<T> result = new ServiceApiResult<T>(code, false, message, 0, null);
        return result;
    }
}
