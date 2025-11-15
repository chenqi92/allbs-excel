package cn.allbs.excel.processor;

import java.lang.reflect.Method;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
public interface NameProcessor {

    /**
     * 解析名称
     *
     * @param args   拦截器对象
     * @param method
     * @param key    表达式
     * @return 解析名称
     */
    String doDetermineName(Object[] args, Method method, String key);
}
