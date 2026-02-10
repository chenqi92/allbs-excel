package cn.allbs.excel.util;

import cn.allbs.excel.annotation.cross.Condition;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.vo.FieldError;
import cn.idev.excel.annotation.ExcelProperty;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.Map;

/**
 * 校验工具类
 * <p>
 * 提供跨字段校验的公共方法
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public final class ValidationHelper {

    private ValidationHelper() {
    }

    /**
     * 获取 Excel 列名
     *
     * @param clazz     类
     * @param fieldName 字段名
     * @return Excel 列名，如果没有 @ExcelProperty 则返回字段名
     */
    public static String getExcelFieldName(Class<?> clazz, String fieldName) {
        Field field = FieldAccessorCache.getField(clazz, fieldName);
        if (field != null) {
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null && excelProperty.value().length > 0) {
                return excelProperty.value()[0];
            }
        }
        return fieldName;
    }

    /**
     * 检查多条件是否满足（AND 关系）
     *
     * @param target     目标对象
     * @param conditions 条件数组
     * @return 所有条件是否都满足
     */
    public static boolean checkMultipleConditions(Object target, Condition[] conditions) {
        for (Condition condition : conditions) {
            Object value = FieldAccessorCache.getFieldValue(target, condition.field());

            if (condition.notEmpty()) {
                if (isEmpty(value)) {
                    return false;
                }
            } else if (condition.isEmpty()) {
                if (isNotEmpty(value)) {
                    return false;
                }

            } else if (condition.value().isEmpty()) {
                // value 为空字符串（默认值），且未设置 notEmpty/isEmpty，视为 notEmpty 检查
                if (isEmpty(value)) {
                    return false;
                }
            } else {
                // 安全的值比较
                if (!safeEquals(condition.value(), value)) {
                    return false;
                }
            }
        }
        return true;
    }

    /**
     * 安全的值比较
     * <p>
     * 正确处理 null 值，避免 String.valueOf(null) 返回 "null" 字符串的问题
     * </p>
     *
     * @param expected 期望值
     * @param actual   实际值
     * @return 是否相等
     */
    public static boolean safeEquals(String expected, Object actual) {
        if (actual == null) {
            return expected == null || expected.isEmpty();
        }
        return expected.equals(String.valueOf(actual));
    }

    /**
     * 检查值是否为空
     * <p>
     * 支持 null、空字符串、空集合、空 Map、空数组
     * </p>
     *
     * @param value 值
     * @return 是否为空
     */
    public static boolean isEmpty(Object value) {
        if (value == null) {
            return true;
        }
        if (value instanceof String) {
            return ((String) value).trim().isEmpty();
        }
        if (value instanceof Collection) {
            return ((Collection<?>) value).isEmpty();
        }
        if (value instanceof Map) {
            return ((Map<?, ?>) value).isEmpty();
        }
        if (value.getClass().isArray()) {
            return java.lang.reflect.Array.getLength(value) == 0;
        }
        return false;
    }

    /**
     * 检查值是否非空
     *
     * @param value 值
     * @return 是否非空
     */
    public static boolean isNotEmpty(Object value) {
        return !isEmpty(value);
    }

    /**
     * 构建字段错误
     *
     * @param clazz      类
     * @param fieldName  字段名
     * @param errorType  错误类型
     * @param message    错误消息
     * @param fieldValue 字段值
     * @return FieldError 对象
     */
    public static FieldError buildFieldError(Class<?> clazz, String fieldName, String errorType,
            String message, Object fieldValue) {
        String excelFieldName = getExcelFieldName(clazz, fieldName);
        return FieldError.builder()
                .fieldName(excelFieldName)
                .propertyName(fieldName)
                .errorType(errorType)
                .message(message)
                .fullMessage("【" + excelFieldName + "】" + message)
                .fieldValue(fieldValue)
                .build();
    }

    /**
     * 构建字段错误（无字段值）
     */
    public static FieldError buildFieldError(Class<?> clazz, String fieldName, String errorType, String message) {
        return buildFieldError(clazz, fieldName, errorType, message, null);
    }

    /**
     * 检查分组是否匹配
     * <p>
     * 如果规则没有指定分组，则匹配所有分组；
     * 如果规则指定了分组，则检查分组是否有继承关系。
     * </p>
     *
     * @param ruleGroups 规则分组
     * @param groups     校验分组
     * @return 是否匹配
     */
    public static boolean matchGroups(Class<?>[] ruleGroups, Class<?>... groups) {
        if (ruleGroups == null || ruleGroups.length == 0) {
            return true; // 没有指定分组，默认所有分组都匹配
        }
        if (groups == null || groups.length == 0) {
            return false; // 规则指定了分组，但校验未指定分组，不匹配
        }
        for (Class<?> ruleGroup : ruleGroups) {
            for (Class<?> group : groups) {
                if (ruleGroup.isAssignableFrom(group)) {
                    return true;
                }
            }
        }
        return false;
    }
}
