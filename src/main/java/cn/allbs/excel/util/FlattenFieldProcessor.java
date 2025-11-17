package cn.allbs.excel.util;

import cn.allbs.excel.annotation.FlattenProperty;
import com.alibaba.excel.annotation.ExcelProperty;
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
        processFields(clazz, "", 0, result, new HashSet<>());

        // 按照顺序排序
        result.sort(Comparator.comparingInt(FlattenFieldInfo::getOrder));

        return result;
    }

    /**
     * 递归处理字段
     *
     * @param clazz         当前类
     * @param pathPrefix    路径前缀（用于递归）
     * @param currentDepth  当前递归深度
     * @param result        结果列表
     * @param processedTypes 已处理的类型（防止循环引用）
     */
    private static void processFields(Class<?> clazz, String pathPrefix, int currentDepth,
                                      List<FlattenFieldInfo> result, Set<Class<?>> processedTypes) {
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
            // 处理普通 @ExcelProperty 字段
            if (field.isAnnotationPresent(ExcelProperty.class) && !field.isAnnotationPresent(FlattenProperty.class)) {
                ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                String[] headNames = excelProperty.value();
                String headName = headNames.length > 0 ? headNames[0] : field.getName();

                FlattenFieldInfo info = new FlattenFieldInfo();
                info.setFieldName(field.getName());
                info.setFieldPath(pathPrefix.isEmpty() ? field.getName() : pathPrefix + "." + field.getName());
                info.setHeadName(headName);
                info.setField(field);
                info.setOrder(baseOrder + excelProperty.index());
                info.setFlatten(false);
                result.add(info);
            }
            // 处理 @FlattenProperty 字段
            else if (field.isAnnotationPresent(FlattenProperty.class)) {
                FlattenProperty flattenProperty = field.getAnnotation(FlattenProperty.class);
                Class<?> fieldType = field.getType();

                // 检查是否超过最大递归深度
                if (flattenProperty.recursive() && currentDepth >= flattenProperty.maxDepth()) {
                    log.warn("Max recursion depth reached for field: {}.{}", clazz.getName(), field.getName());
                    continue;
                }

                // 扫描嵌套类的字段
                List<FlattenFieldInfo> nestedFields = new ArrayList<>();
                Set<Class<?>> nestedProcessedTypes = new HashSet<>(processedTypes);

                if (flattenProperty.recursive()) {
                    processFields(fieldType, field.getName(), currentDepth + 1, nestedFields, nestedProcessedTypes);
                } else {
                    // 非递归模式，只展开一层
                    Field[] nestedFieldArray = fieldType.getDeclaredFields();
                    int nestedBaseOrder = 0;

                    for (Field nestedField : nestedFieldArray) {
                        if (nestedField.isAnnotationPresent(ExcelProperty.class)) {
                            ExcelProperty excelProperty = nestedField.getAnnotation(ExcelProperty.class);
                            String[] headNames = excelProperty.value();
                            String headName = headNames.length > 0 ? headNames[0] : nestedField.getName();

                            // 应用前缀和后缀
                            String finalHeadName = flattenProperty.prefix() + headName + flattenProperty.suffix();

                            FlattenFieldInfo info = new FlattenFieldInfo();
                            info.setFieldName(nestedField.getName());
                            info.setFieldPath(field.getName() + "." + nestedField.getName());
                            info.setHeadName(finalHeadName);
                            info.setField(nestedField);
                            info.setParentField(field);
                            info.setOrder(baseOrder + flattenProperty.orderOffset() + nestedBaseOrder + excelProperty.index());
                            info.setFlatten(true);
                            info.setFlattenAnnotation(flattenProperty);
                            nestedFields.add(info);
                            nestedBaseOrder++;
                        }
                    }
                }

                result.addAll(nestedFields);
            }
        }

        processedTypes.remove(clazz);
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
