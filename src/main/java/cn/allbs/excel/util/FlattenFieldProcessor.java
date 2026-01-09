package cn.allbs.excel.util;

import cn.allbs.excel.annotation.FlattenProperty;
import cn.idev.excel.annotation.ExcelProperty;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * 嵌套对象字段展开处理器
 * <p>
 * 用于处理 @FlattenProperty 注解，自动展开嵌套对象中的字段
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class FlattenFieldProcessor {

    /**
     * 扫描类并生成展开后的字段信息
     *
     * @param clazz 要扫描的类
     * @return 展开后的字段信息列表
     */
    public static List<FlattenFieldInfo> processFlattenFields(Class<?> clazz) {
        List<FlattenFieldInfo> result = new ArrayList<>();
        processFields(clazz, "", "", "", 0, result, new HashSet<>(), null);

        // 按照顺序排序
        result.sort(Comparator.comparingInt(FlattenFieldInfo::getOrder));

        return result;
    }

    /**
     * 递归处理字段
     *
     * @param clazz         当前类
     * @param pathPrefix    路径前缀（用于递归）
     * @param headPrefix    表头前缀
     * @param headSuffix    表头后缀
     * @param currentDepth  当前递归深度
     * @param result        结果列表
     * @param processedTypes 已处理的类型（防止循环引用）
     * @param parentField   父字段（用于展开字段）
     */
    private static void processFields(Class<?> clazz, String pathPrefix, String headPrefix, String headSuffix,
                                      int currentDepth, List<FlattenFieldInfo> result, Set<Class<?>> processedTypes,
                                      Field parentField) {
        if (clazz == null || clazz == Object.class) {
            return;
        }

        // 防止循环引用
        if (processedTypes.contains(clazz)) {
            log.warn("Circular reference detected for class: {}", clazz.getName());
            return;
        }
        processedTypes.add(clazz);

        Field[] fields = clazz.getDeclaredFields();
        int baseOrder = result.size();

        for (Field field : fields) {
            // 处理 @FlattenProperty 字段（仅在顶层 currentDepth == 0）
            if (field.isAnnotationPresent(FlattenProperty.class) && currentDepth == 0) {
                FlattenProperty flattenProperty = field.getAnnotation(FlattenProperty.class);
                Class<?> fieldType = field.getType();

                // 检查是否超过最大递归深度
                if (currentDepth >= flattenProperty.maxDepth()) {
                    log.warn("Max recursion depth reached for field: {}.{}", clazz.getName(), field.getName());
                    continue;
                }

                // 扫描嵌套类的字段
                List<FlattenFieldInfo> nestedFields = new ArrayList<>();
                Set<Class<?>> nestedProcessedTypes = new HashSet<>(processedTypes);

                // 合并前缀和后缀
                String newPrefix = headPrefix + flattenProperty.prefix();
                String newSuffix = flattenProperty.suffix() + headSuffix;
                String newPath = pathPrefix.isEmpty() ? field.getName() : pathPrefix + "." + field.getName();

                // 根据是否递归，传递不同的深度和模式
                if (flattenProperty.recursive()) {
                    // 递归模式：继续展开嵌套对象类型字段
                    processFieldsRecursive(fieldType, newPath, newPrefix, newSuffix,
                                          currentDepth + 1, flattenProperty.maxDepth(),
                                          nestedFields, nestedProcessedTypes, field);
                } else {
                    // 非递归模式：只展开一层
                    processFields(fieldType, newPath, newPrefix, newSuffix,
                                 currentDepth + 1, nestedFields, nestedProcessedTypes, field);
                }

                // 应用 orderOffset
                for (FlattenFieldInfo info : nestedFields) {
                    info.setOrder(info.getOrder() + flattenProperty.orderOffset());
                    info.setFlattenAnnotation(flattenProperty);
                }

                result.addAll(nestedFields);
            }
            // 只处理有 @ExcelProperty 的字段（非 @FlattenProperty）
            else if (field.isAnnotationPresent(ExcelProperty.class) && !field.isAnnotationPresent(FlattenProperty.class)) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                String[] headNames = excelProperty.value();
                String headName = headNames.length > 0 ? headNames[0] : field.getName();

                // 应用前缀和后缀（如果在嵌套处理中）
                String finalHeadName = headPrefix + headName + headSuffix;

                FlattenFieldInfo info = new FlattenFieldInfo();
                info.setFieldName(field.getName());
                info.setFieldPath(pathPrefix.isEmpty() ? field.getName() : pathPrefix + "." + field.getName());
                info.setHeadName(finalHeadName);
                info.setField(field);
                info.setParentField(parentField);
                info.setOrder(baseOrder + excelProperty.index());
                info.setFlatten(parentField != null);  // 如果有父字段，说明是展开字段
                result.add(info);
            }
        }

        processedTypes.remove(clazz);
    }

    /**
     * 递归模式处理字段（自动展开嵌套对象）
     *
     * @param clazz         当前类
     * @param pathPrefix    路径前缀
     * @param headPrefix    表头前缀
     * @param headSuffix    表头后缀
     * @param currentDepth  当前递归深度
     * @param maxDepth      最大递归深度
     * @param result        结果列表
     * @param processedTypes 已处理的类型（防止循环引用）
     * @param parentField   父字段
     */
    private static void processFieldsRecursive(Class<?> clazz, String pathPrefix, String headPrefix, String headSuffix,
                                               int currentDepth, int maxDepth, List<FlattenFieldInfo> result,
                                               Set<Class<?>> processedTypes, Field parentField) {
        if (clazz == null || clazz == Object.class) {
            return;
        }

        // 防止循环引用
        if (processedTypes.contains(clazz)) {
            log.warn("Circular reference detected for class: {}", clazz.getName());
            return;
        }

        // 检查递归深度
        if (currentDepth >= maxDepth) {
            log.warn("Max recursion depth {} reached for class: {}", maxDepth, clazz.getName());
            return;
        }

        processedTypes.add(clazz);

        Field[] fields = clazz.getDeclaredFields();
        int baseOrder = result.size();

        for (Field field : fields) {
            // 处理有 @ExcelProperty 的字段
            if (field.isAnnotationPresent(ExcelProperty.class)) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                String[] headNames = excelProperty.value();
                String headName = headNames.length > 0 ? headNames[0] : field.getName();

                // 应用前缀和后缀
                String finalHeadName = headPrefix + headName + headSuffix;

                FlattenFieldInfo info = new FlattenFieldInfo();
                info.setFieldName(field.getName());
                info.setFieldPath(pathPrefix.isEmpty() ? field.getName() : pathPrefix + "." + field.getName());
                info.setHeadName(finalHeadName);
                info.setField(field);
                info.setParentField(parentField);
                info.setOrder(baseOrder + excelProperty.index());
                info.setFlatten(true);
                result.add(info);
            }
            // 递归模式：自动展开对象类型字段（即使没有 @ExcelProperty 或 @FlattenProperty）
            else if (isComplexType(field.getType())) {
                Class<?> fieldType = field.getType();
                String newPath = pathPrefix.isEmpty() ? field.getName() : pathPrefix + "." + field.getName();

                // 继续递归处理（保持当前前缀和后缀）
                List<FlattenFieldInfo> nestedFields = new ArrayList<>();
                Set<Class<?>> nestedProcessedTypes = new HashSet<>(processedTypes);

                processFieldsRecursive(fieldType, newPath, headPrefix, headSuffix,
                                      currentDepth + 1, maxDepth, nestedFields, nestedProcessedTypes, parentField);

                result.addAll(nestedFields);
            }
        }

        processedTypes.remove(clazz);
    }

    /**
     * 判断是否为复杂类型（需要递归展开的类型）
     *
     * @param clazz 类型
     * @return true 如果是复杂类型
     */
    private static boolean isComplexType(Class<?> clazz) {
        if (clazz == null || clazz == Object.class) {
            return false;
        }

        // 排除基本类型和包装类型
        if (clazz.isPrimitive() ||
            clazz == String.class ||
            clazz == Integer.class ||
            clazz == Long.class ||
            clazz == Double.class ||
            clazz == Float.class ||
            clazz == Boolean.class ||
            clazz == Short.class ||
            clazz == Byte.class ||
            clazz == Character.class ||
            Number.class.isAssignableFrom(clazz)) {
            return false;
        }

        // 排除日期时间类型
        if (java.util.Date.class.isAssignableFrom(clazz) ||
            java.time.temporal.Temporal.class.isAssignableFrom(clazz)) {
            return false;
        }

        // 排除集合、数组、Map 等
        if (clazz.isArray() ||
            java.util.Collection.class.isAssignableFrom(clazz) ||
            java.util.Map.class.isAssignableFrom(clazz)) {
            return false;
        }

        // 排除枚举
        if (clazz.isEnum()) {
            return false;
        }

        // 其他都视为复杂类型（自定义类）
        return true;
    }

    /**
     * 从对象中提取展开字段的值
     *
     * @param obj   源对象
     * @param info  字段信息
     * @return 字段值
     */
    public static Object extractValue(Object obj, FlattenFieldInfo info) {
        if (obj == null) {
            return handleNullValue(info);
        }

        try {
            // 如果是展开字段，需要先获取父对象
            if (info.isFlatten() && info.getParentField() != null) {
                Object parentObj = getFieldValue(obj, info.getParentField());
                if (parentObj == null) {
                    return handleNullValue(info);
                }
                return getFieldValue(parentObj, info.getField());
            } else {
                // 普通字段或多层嵌套路径
                return resolveValueByPath(obj, info.getFieldPath());
            }
        } catch (Exception e) {
            log.error("Failed to extract value for field: {}", info.getFieldPath(), e);
            return handleNullValue(info);
        }
    }

    /**
     * 根据路径解析值
     *
     * @param obj  源对象
     * @param path 字段路径
     * @return 字段值
     */
    private static Object resolveValueByPath(Object obj, String path) throws Exception {
        if (obj == null || path == null || path.isEmpty()) {
            return null;
        }

        String[] pathParts = path.split("\\.");
        Object current = obj;

        for (String part : pathParts) {
            if (current == null) {
                return null;
            }
            current = getFieldValueByName(current, part);
        }

        return current;
    }

    /**
     * 获取字段值
     *
     * @param obj   对象
     * @param field 字段
     * @return 字段值
     * @throws Exception 反射异常
     */
    private static Object getFieldValue(Object obj, Field field) throws Exception {
        // 优先使用 getter 方法
        try {
            String getterName = "get" + capitalize(field.getName());
            Method getter = obj.getClass().getMethod(getterName);
            return getter.invoke(obj);
        } catch (NoSuchMethodException e) {
            // 尝试 is 开头的 getter
            try {
                String isGetterName = "is" + capitalize(field.getName());
                Method isGetter = obj.getClass().getMethod(isGetterName);
                return isGetter.invoke(obj);
            } catch (NoSuchMethodException ex) {
                // 直接访问字段
                field.setAccessible(true);
                return field.get(obj);
            }
        }
    }

    /**
     * 根据字段名获取值
     *
     * @param obj       对象
     * @param fieldName 字段名
     * @return 字段值
     * @throws Exception 反射异常
     */
    private static Object getFieldValueByName(Object obj, String fieldName) throws Exception {
        Field field = findField(obj.getClass(), fieldName);
        if (field == null) {
            throw new NoSuchFieldException("Field not found: " + fieldName);
        }
        return getFieldValue(obj, field);
    }

    /**
     * 查找字段（包括父类）
     *
     * @param clazz     类
     * @param fieldName 字段名
     * @return 字段对象
     */
    private static Field findField(Class<?> clazz, String fieldName) {
        Class<?> currentClass = clazz;
        while (currentClass != null && currentClass != Object.class) {
            try {
                return currentClass.getDeclaredField(fieldName);
            } catch (NoSuchFieldException e) {
                currentClass = currentClass.getSuperclass();
            }
        }
        return null;
    }

    /**
     * 处理 null 值
     *
     * @param info 字段信息
     * @return null 值的处理结果
     */
    private static Object handleNullValue(FlattenFieldInfo info) {
        if (!info.isFlatten() || info.getFlattenAnnotation() == null) {
            return "";
        }

        FlattenProperty annotation = info.getFlattenAnnotation();
        switch (annotation.nullStrategy()) {
            case SKIP:
                return null;
            case CUSTOM:
                return annotation.nullValue();
            case EMPTY_STRING:
            default:
                return "";
        }
    }

    /**
     * 首字母大写
     *
     * @param str 字符串
     * @return 首字母大写的字符串
     */
    private static String capitalize(String str) {
        if (str == null || str.isEmpty()) {
            return str;
        }
        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }

    /**
     * 展开字段信息
     */
    @Data
    public static class FlattenFieldInfo {
        /**
         * 字段名
         */
        private String fieldName;

        /**
         * 字段路径（用于多层嵌套）
         */
        private String fieldPath;

        /**
         * 表头名称
         */
        private String headName;

        /**
         * 字段对象
         */
        private Field field;

        /**
         * 父字段（对于展开字段）
         */
        private Field parentField;

        /**
         * 字段顺序
         */
        private int order;

        /**
         * 是否是展开字段
         */
        private boolean flatten;

        /**
         * FlattenProperty 注解（如果是展开字段）
         */
        private FlattenProperty flattenAnnotation;
    }
}
