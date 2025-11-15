package cn.allbs.excel.handle;


import cn.allbs.excel.annotation.ExportExcel;

import jakarta.servlet.http.HttpServletResponse;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
public interface SheetWriteHandler {

    /**
     * 是否支持
     *
     * @param obj 返回对象
     * @return boolean
     */
    boolean support(Object obj);

    /**
     * 校验
     *
     * @param responseExcel 注解
     */
    void check(ExportExcel responseExcel);

    /**
     * 返回的对象
     *
     * @param o             obj
     * @param response      输出对象
     * @param responseExcel 注解
     */
    void export(Object o, HttpServletResponse response, ExportExcel responseExcel);

    /**
     * 写成对象
     *
     * @param o             obj
     * @param response      输出对象
     * @param responseExcel 注解
     */
    void write(Object o, HttpServletResponse response, ExportExcel responseExcel);
}
