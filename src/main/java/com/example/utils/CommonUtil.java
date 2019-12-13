package com.example.utils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class CommonUtil {

    /**
     * 分割-逗号
     */
    public static final String COMMON_SPLIT = ",";

    /**
     * 判断对象是否为null和空
     *
     * @param object 对象
     * @return boolean
     */
    @SuppressWarnings("rawtypes")
    public static boolean isObjectNull(Object object) {
        if (object == null) {
            return true;
        } else if (object instanceof java.util.Collection) {
            return ((java.util.Collection) object).isEmpty() ? true : false;
        } else if (object instanceof java.util.Map) {
            return ((java.util.Map) object).isEmpty() ? true : false;
        } else if (object instanceof String) {
            return ((String) object).trim().length() == 0 ? true : false;
        }
        return false;
    }

    /*
     * 判断字符串是否可用
     * */
    public static boolean isValidStr(String str) {
        boolean isValid = true;

        if (null == str) {
            isValid = false;
        }

        if ("".equals(str)) {
            isValid = false;
        }

        if (null != str && str.toLowerCase().trim().equals("null")) {
            isValid = false;
        }

        return isValid;
    }


    public static List<String> splitStr(String str, String separator) {
        List<String> list = new ArrayList<String>();
        if (CommonUtil.isObjectNull(str)) {
            return null;
        }
        String[] split = str.split(separator);
        list.addAll(Arrays.asList(split));
        return list;
    }


    public static void main(String[] args) {

    }

}