package cn.allbs.excel.convert;

import cn.allbs.excel.annotation.NestedProperty;
import cn.allbs.excel.util.NestedFieldResolver;
import cn.idev.excel.converters.Converter;
import cn.idev.excel.enums.CellDataTypeEnum;
import cn.idev.excel.metadata.GlobalConfiguration;
import cn.idev.excel.metadata.data.ReadCellData;
import cn.idev.excel.metadata.data.WriteCellData;
import cn.idev.excel.metadata.property.ExcelContentProperty;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 嵌套对象转换器
 * <p>
 * 用于将嵌套对象的字段值提取并导出到 Excel
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty(value = "部门名称", converter = NestedObjectConverter.class)
 * &#64;NestedProperty("dept.name")
 * private Department dept;
 *
 * &#64;ExcelProperty(value = "领导姓名", converter = NestedObjectConverter.class)
 * &#64;NestedProperty(value = "dept.leader.name", nullValue = "暂无")
 * private Department dept;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class NestedObjectConverter implements Converter<Object> {

    /**
     * 路径片段模式：支持 fieldName、fieldName[index]、fieldName[key]、fieldName[*]
     */
    private static final Pattern PATH_SEGMENT_PATTERN = Pattern.compile("^([^\\[]+)(?:\\[([^\\]]+)\\])?$");

    /**
     * 直接访问模式：[index]、[key]、[*]
     */
    private static final Pattern DIRECT_ACCESS_PATTERN = Pattern.compile("^\\[([^\\]]+)\\]$");

    @Override
    public Class<?> supportJavaTypeKey() {
        return Object.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 导出时：将嵌套对象转换为 Excel 单元格数据
     * <p>
     * 根据 NestedProperty 注解中的路径表达式提取嵌套字段的值
     * </p>
     */
    @Override
    public WriteCellData<?> convertToExcelData(Object value, ExcelContentProperty contentProperty,
                                                GlobalConfiguration globalConfiguration) {
        if (value == null) {
            return new WriteCellData<>("");
        }

        try {
            // 获取字段
            Field field = contentProperty.getField();
            if (field == null) {
                log.warn("Field is null, returning original value");
                return new WriteCellData<>(value.toString());
            }

            // 获取 NestedProperty 注解
            NestedProperty nestedProperty = field.getAnnotation(NestedProperty.class);
            if (nestedProperty == null) {
                // 如果没有 NestedProperty 注解，直接返回对象的 toString()
                log.debug("No NestedProperty annotation found on field: {}, using toString()", field.getName());
                return new WriteCellData<>(value.toString());
            }

            // 根据注解配置提取嵌套值
            Object nestedValue = NestedFieldResolver.resolveNestedValue(value, nestedProperty);
            String cellValue = nestedValue != null ? nestedValue.toString() : nestedProperty.nullValue();

            return new WriteCellData<>(cellValue);

        } catch (Exception e) {
            log.error("Failed to convert nested object to excel data", e);
            return new WriteCellData<>(value.toString());
        }
    }

    /**
     * 导入时：将 Excel 单元格数据转换为嵌套对象
     * <p>
     * 根据 NestedProperty 注解中的路径表达式，反向构建嵌套对象
     * </p>
     */
    @Override
    public Object convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
                                     GlobalConfiguration globalConfiguration) {
        try {
            Field field = contentProperty.getField();
            if (field == null) {
                log.warn("Field is null, returning string value");
                return cellData.getStringValue();
            }

            NestedProperty nestedProperty = field.getAnnotation(NestedProperty.class);
            String cellValue = cellData.getStringValue();

            // 处理空值
            if (cellValue == null || cellValue.isEmpty()) {
                return null;
            }

            // 如果配置了 nullValue，且单元格值等于 nullValue，返回 null
            if (nestedProperty != null && cellValue.equals(nestedProperty.nullValue())) {
                return null;
            }

            Class<?> fieldType = field.getType();

            // 如果没有 NestedProperty 注解，尝试简单转换
            if (nestedProperty == null) {
                return convertSimpleValue(cellValue, fieldType);
            }

            String path = nestedProperty.value();

            // 根据字段类型和路径构建对象
            if (List.class.isAssignableFrom(fieldType)) {
                return createListWithValue(field, path, cellValue, nestedProperty);
            } else if (Map.class.isAssignableFrom(fieldType)) {
                return createMapWithValue(field, path, cellValue);
            } else {
                // 普通对象类型
                return createNestedObject(fieldType, path, cellValue);
            }

        } catch (Exception e) {
            log.error("Failed to convert excel data to nested object", e);
            return null;
        }
    }

    /**
     * 创建嵌套对象并设置值
     */
    private Object createNestedObject(Class<?> targetType, String path, String value) throws Exception {
        if (path == null || path.isEmpty()) {
            return convertSimpleValue(value, targetType);
        }

        // 创建根对象
        Object rootObject = targetType.getDeclaredConstructor().newInstance();

        // 设置嵌套值
        setNestedValue(rootObject, path, value);

        return rootObject;
    }

    /**
     * 根据路径设置嵌套值
     */
    private void setNestedValue(Object obj, String path, String value) throws Exception {
        if (obj == null || path == null || path.isEmpty()) {
            return;
        }

        // 检查是否是直接访问模式（如 [0]、[*]、[city]）
        Matcher directMatcher = DIRECT_ACCESS_PATTERN.matcher(path.trim());
        if (directMatcher.matches()) {
            // 直接访问模式不适用于普通对象，跳过
            log.warn("Direct access pattern not supported for import: {}", path);
            return;
        }

        // 分割路径
        String[] parts = path.split("\\.", 2);
        String currentField = parts[0];

        Matcher segmentMatcher = PATH_SEGMENT_PATTERN.matcher(currentField.trim());
        if (!segmentMatcher.matches()) {
            log.warn("Invalid path segment: {}", currentField);
            return;
        }

        String fieldName = segmentMatcher.group(1);
        String accessor = segmentMatcher.group(2); // null, index, key, or *

        Field field = findField(obj.getClass(), fieldName);
        if (field == null) {
            log.warn("Field {} not found in {}", fieldName, obj.getClass().getName());
            return;
        }

        if (parts.length == 1 && accessor == null) {
            // 最后一层，直接设置值
            Object convertedValue = convertSimpleValue(value, field.getType());
            setFieldValue(obj, field, convertedValue);
        } else if (parts.length == 1 && accessor != null) {
            // 最后一层但有访问器（如 list[0]、map[key]）
            handleAccessorField(obj, field, accessor, value);
        } else {
            // 还有嵌套层级，需要创建中间对象
            Object nestedObj = getFieldValue(obj, field);
            if (nestedObj == null) {
                nestedObj = field.getType().getDeclaredConstructor().newInstance();
                setFieldValue(obj, field, nestedObj);
            }
            setNestedValue(nestedObj, parts[1], value);
        }
    }

    /**
     * 处理带访问器的字段（如 list[0]、map[key]）
     */
    @SuppressWarnings("unchecked")
    private void handleAccessorField(Object obj, Field field, String accessor, String value) throws Exception {
        Class<?> fieldType = field.getType();

        if (List.class.isAssignableFrom(fieldType)) {
            int index = Integer.parseInt(accessor);
            List<Object> list = (List<Object>) getFieldValue(obj, field);
            if (list == null) {
                list = new ArrayList<>();
                setFieldValue(obj, field, list);
            }
            // 确保 list 大小足够
            while (list.size() <= index) {
                list.add(null);
            }
            // 获取 List 元素类型
            Type genericType = field.getGenericType();
            Class<?> elementType = getListElementType(genericType);
            list.set(index, convertSimpleValue(value, elementType));
        } else if (Map.class.isAssignableFrom(fieldType)) {
            Map<String, Object> map = (Map<String, Object>) getFieldValue(obj, field);
            if (map == null) {
                map = new HashMap<>();
                setFieldValue(obj, field, map);
            }
            map.put(accessor, value);
        }
    }

    /**
     * 创建 List 并设置值
     */
    private Object createListWithValue(Field field, String path, String value, NestedProperty annotation) throws Exception {
        List<Object> list = new ArrayList<>();

        // 检查路径格式
        Matcher directMatcher = DIRECT_ACCESS_PATTERN.matcher(path.trim());
        if (directMatcher.matches()) {
            String accessor = directMatcher.group(1);
            if ("*".equals(accessor)) {
                // [*] 模式：根据分隔符拆分值
                String separator = annotation.separator();
                String[] values = value.split(Pattern.quote(separator));
                Type genericType = field.getGenericType();
                Class<?> elementType = getListElementType(genericType);
                for (String v : values) {
                    String trimmed = v.trim();
                    if (!trimmed.isEmpty()) {
                        list.add(convertSimpleValue(trimmed, elementType));
                    }
                }
            } else if (accessor.matches("\\d+")) {
                // [index] 模式：设置指定索引的值
                int index = Integer.parseInt(accessor);
                Type genericType = field.getGenericType();
                Class<?> elementType = getListElementType(genericType);
                while (list.size() <= index) {
                    list.add(null);
                }
                list.set(index, convertSimpleValue(value, elementType));
            }
        }

        return list;
    }

    /**
     * 创建 Map 并设置值
     */
    private Object createMapWithValue(Field field, String path, String value) throws Exception {
        Map<String, Object> map = new HashMap<>();

        Matcher directMatcher = DIRECT_ACCESS_PATTERN.matcher(path.trim());
        if (directMatcher.matches()) {
            String key = directMatcher.group(1);
            map.put(key, value);
        }

        return map;
    }

    /**
     * 获取 List 的元素类型
     */
    private Class<?> getListElementType(Type genericType) {
        if (genericType instanceof ParameterizedType) {
            Type[] typeArgs = ((ParameterizedType) genericType).getActualTypeArguments();
            if (typeArgs.length > 0 && typeArgs[0] instanceof Class) {
                return (Class<?>) typeArgs[0];
            }
        }
        return String.class;
    }

    /**
     * 简单值类型转换
     */
    private Object convertSimpleValue(String value, Class<?> targetType) {
        if (value == null) return null;
        if (targetType == String.class) return value;
        else if (targetType == Integer.class || targetType == int.class) return Integer.parseInt(value);
        else if (targetType == Long.class || targetType == long.class) return Long.parseLong(value);
        else if (targetType == Double.class || targetType == double.class) return Double.parseDouble(value);
        else if (targetType == Float.class || targetType == float.class) return Float.parseFloat(value);
        else if (targetType == Boolean.class || targetType == boolean.class) return Boolean.parseBoolean(value);
        else if (targetType == Short.class || targetType == short.class) return Short.parseShort(value);
        else if (targetType == Byte.class || targetType == byte.class) return Byte.parseByte(value);
        else return value;
    }

    /**
     * 查找字段（包括父类）
     */
    private Field findField(Class<?> clazz, String fieldName) {
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
     * 设置字段值
     */
    private void setFieldValue(Object obj, Field field, Object value) throws Exception {
        // 优先使用 setter 方法
        String setterName = "set" + capitalize(field.getName());
        try {
            Method setter = obj.getClass().getMethod(setterName, field.getType());
            setter.invoke(obj, value);
        } catch (NoSuchMethodException e) {
            // 没有 setter，直接设置字段
            field.setAccessible(true);
            field.set(obj, value);
        }
    }

    /**
     * 获取字段值
     */
    private Object getFieldValue(Object obj, Field field) throws Exception {
        String getterName = "get" + capitalize(field.getName());
        try {
            Method getter = obj.getClass().getMethod(getterName);
            return getter.invoke(obj);
        } catch (NoSuchMethodException e) {
            // 尝试 is 开头的 getter
            try {
                String isGetterName = "is" + capitalize(field.getName());
                Method isGetter = obj.getClass().getMethod(isGetterName);
                return isGetter.invoke(obj);
            } catch (NoSuchMethodException ex) {
                field.setAccessible(true);
                return field.get(obj);
            }
        }
    }

    /**
     * 首字母大写
     */
    private String capitalize(String str) {
        if (str == null || str.isEmpty()) return str;
        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }
}
