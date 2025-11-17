package cn.allbs.excel.util;

import cn.allbs.excel.annotation.NestedProperty;
import lombok.extern.slf4j.Slf4j;
import org.springframework.util.StringUtils;

import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 嵌套字段解析工具类
 * <p>
 * 用于解析嵌套对象的字段路径，支持对象、集合、Map、数组等复杂类型
 * </p>
 *
 * <p>支持的路径语法：</p>
 * <ul>
 *   <li>对象字段：dept.name</li>
 *   <li>集合索引：depts[0].name</li>
 *   <li>集合全部：depts[*].name</li>
 *   <li>Map键访问：properties[city]</li>
 *   <li>数组索引：tags[0]</li>
 * </ul>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class NestedFieldResolver {

    /**
     * 路径片段模式：支持 fieldName、fieldName[index]、fieldName[key]、fieldName[*]
     */
    private static final Pattern PATH_SEGMENT_PATTERN = Pattern.compile("^([^\\[]+)(?:\\[([^\\]]+)\\])?$");

    /**
     * 根据嵌套路径从对象中提取值
     *
     * @param obj  源对象
     * @param path 嵌套路径
     * @return 提取到的值，如果为 null 则返回 null
     */
    public static Object resolveNestedValue(Object obj, String path) {
        return resolveNestedValue(obj, path, "", ",", 0);
    }

    /**
     * 根据嵌套路径从对象中提取值
     *
     * @param obj          源对象
     * @param path         嵌套路径
     * @param defaultValue 默认值
     * @return 提取到的值，如果为 null 则返回默认值
     */
    public static Object resolveNestedValue(Object obj, String path, String defaultValue) {
        return resolveNestedValue(obj, path, defaultValue, ",", 0);
    }

    /**
     * 根据嵌套路径从对象中提取值
     *
     * @param obj          源对象
     * @param path         嵌套路径
     * @param defaultValue 默认值
     * @param separator    集合拼接分隔符
     * @param maxJoinSize  集合最大拼接数量
     * @return 提取到的值，如果为 null 则返回默认值
     */
    public static Object resolveNestedValue(Object obj, String path, String defaultValue,
                                           String separator, int maxJoinSize) {
        if (obj == null || !StringUtils.hasText(path)) {
            return defaultValue;
        }

        try {
            // 分割路径片段
            List<PathSegment> segments = parsePathSegments(path);
            Object currentObj = obj;

            // 逐层访问
            for (int i = 0; i < segments.size(); i++) {
                PathSegment segment = segments.get(i);
                if (currentObj == null) {
                    return defaultValue;
                }

                // 检查是否是 COLLECTION_ALL，且后面还有路径片段
                if (segment.accessorType == AccessorType.COLLECTION_ALL && i < segments.size() - 1) {
                    // 对集合每个元素继续提取后续路径
                    List<PathSegment> remainingSegments = segments.subList(i + 1, segments.size());
                    currentObj = resolveCollectionWithPath(currentObj, segment, remainingSegments,
                                                          separator, maxJoinSize);
                    break; // 已处理完所有片段
                } else {
                    currentObj = resolveSegment(currentObj, segment, separator, maxJoinSize);
                }
            }

            return currentObj != null ? currentObj : defaultValue;
        } catch (Exception e) {
            log.warn("Failed to resolve nested value from path: {}, error: {}", path, e.getMessage());
            return defaultValue;
        }
    }

    /**
     * 根据 NestedProperty 注解从对象中提取值
     *
     * @param obj        源对象
     * @param annotation NestedProperty 注解
     * @return 提取到的值
     */
    public static Object resolveNestedValue(Object obj, NestedProperty annotation) {
        if (annotation == null) {
            return obj;
        }

        try {
            Object value = resolveNestedValue(obj, annotation.value(), annotation.nullValue(),
                                             annotation.separator(), annotation.maxJoinSize());
            return value;
        } catch (Exception e) {
            if (annotation.ignoreException()) {
                log.warn("Failed to resolve nested value from path: {}, using null value: {}",
                        annotation.value(), annotation.nullValue());
                return annotation.nullValue();
            } else {
                throw new RuntimeException("Failed to resolve nested property: " + annotation.value(), e);
            }
        }
    }

    /**
     * 解析路径为片段列表
     *
     * @param path 路径字符串，如 "dept.leaders[0].name"
     * @return 路径片段列表
     */
    private static List<PathSegment> parsePathSegments(String path) {
        List<PathSegment> segments = new ArrayList<>();
        String[] parts = path.split("\\.");

        for (String part : parts) {
            Matcher matcher = PATH_SEGMENT_PATTERN.matcher(part.trim());
            if (matcher.matches()) {
                String fieldName = matcher.group(1);
                String accessor = matcher.group(2); // null, index, key, or *

                if (accessor == null) {
                    // 简单字段访问：dept
                    segments.add(new PathSegment(fieldName, AccessorType.FIELD, null));
                } else if ("*".equals(accessor)) {
                    // 集合全部访问：depts[*]
                    segments.add(new PathSegment(fieldName, AccessorType.COLLECTION_ALL, null));
                } else if (accessor.matches("\\d+")) {
                    // 索引访问：depts[0]
                    segments.add(new PathSegment(fieldName, AccessorType.INDEX, Integer.parseInt(accessor)));
                } else {
                    // Map键访问：properties[city]
                    segments.add(new PathSegment(fieldName, AccessorType.MAP_KEY, accessor));
                }
            } else {
                throw new IllegalArgumentException("Invalid path segment: " + part);
            }
        }

        return segments;
    }

    /**
     * 解析单个路径片段
     *
     * @param obj         当前对象
     * @param segment     路径片段
     * @param separator   分隔符
     * @param maxJoinSize 最大拼接数量
     * @return 解析后的值
     * @throws Exception 解析异常
     */
    private static Object resolveSegment(Object obj, PathSegment segment, String separator, int maxJoinSize)
            throws Exception {
        if (obj == null) {
            return null;
        }

        // 首先获取字段值
        Object fieldValue = getFieldValue(obj, segment.fieldName);
        if (fieldValue == null) {
            return null;
        }

        // 根据访问器类型处理
        switch (segment.accessorType) {
            case FIELD:
                // 简单字段访问
                return fieldValue;

            case INDEX:
                // 索引访问：支持 List、Set、数组
                return getIndexedValue(fieldValue, (Integer) segment.accessor);

            case MAP_KEY:
                // Map 键访问
                return getMapValue(fieldValue, (String) segment.accessor);

            case COLLECTION_ALL:
                // 集合全部访问，拼接结果
                return joinCollectionValues(fieldValue, separator, maxJoinSize);

            default:
                return fieldValue;
        }
    }

    /**
     * 处理集合类型且后续还有路径片段的情况
     * 例如：depts[*].name - 提取每个 dept 的 name 并拼接
     *
     * @param obj               当前对象
     * @param segment           当前片段（必须是 COLLECTION_ALL 类型）
     * @param remainingSegments 剩余路径片段
     * @param separator         拼接分隔符
     * @param maxJoinSize       最大拼接数量
     * @return 拼接后的字符串
     * @throws Exception 解析异常
     */
    private static Object resolveCollectionWithPath(Object obj, PathSegment segment,
                                                    List<PathSegment> remainingSegments,
                                                    String separator, int maxJoinSize) throws Exception {
        // 获取集合字段
        Object fieldValue = getFieldValue(obj, segment.fieldName);
        if (fieldValue == null) {
            return null;
        }

        Collection<?> collection = toCollection(fieldValue);
        if (collection == null || collection.isEmpty()) {
            return null;
        }

        // 限制数量
        if (maxJoinSize > 0 && collection.size() > maxJoinSize) {
            collection = collection.stream().limit(maxJoinSize).collect(Collectors.toList());
        }

        // 对每个元素继续提取后续路径
        List<String> results = new ArrayList<>();
        for (Object element : collection) {
            if (element == null) {
                continue;
            }

            // 继续解析剩余路径
            Object value = element;
            for (PathSegment seg : remainingSegments) {
                if (value == null) {
                    break;
                }
                value = resolveSegment(value, seg, separator, maxJoinSize);
            }

            if (value != null) {
                results.add(value.toString());
            }
        }

        return String.join(separator, results);
    }

    /**
     * 将对象转换为 Collection
     *
     * @param obj 对象
     * @return Collection 对象，如果无法转换则返回 null
     */
    private static Collection<?> toCollection(Object obj) {
        if (obj == null) {
            return null;
        }

        // 数组转集合
        if (obj.getClass().isArray()) {
            int length = Array.getLength(obj);
            List<Object> list = new ArrayList<>(length);
            for (int i = 0; i < length; i++) {
                list.add(Array.get(obj, i));
            }
            return list;
        }

        // 集合类型
        if (obj instanceof Collection) {
            return (Collection<?>) obj;
        }

        return null;
    }

    /**
     * 获取对象指定字段的值
     *
     * @param obj       对象
     * @param fieldName 字段名
     * @return 字段值
     * @throws Exception 反射异常
     */
    private static Object getFieldValue(Object obj, String fieldName) throws Exception {
        if (obj == null) {
            return null;
        }

        // 特殊处理：如果对象是 Map，直接按 key 获取
        if (obj instanceof Map) {
            return ((Map<?, ?>) obj).get(fieldName);
        }

        Class<?> clazz = obj.getClass();

        // 优先尝试使用 getter 方法
        try {
            String getterName = "get" + capitalize(fieldName);
            Method getter = clazz.getMethod(getterName);
            return getter.invoke(obj);
        } catch (NoSuchMethodException e) {
            // 尝试 is 开头的 getter (用于 boolean 类型)
            try {
                String isGetterName = "is" + capitalize(fieldName);
                Method isGetter = clazz.getMethod(isGetterName);
                return isGetter.invoke(obj);
            } catch (NoSuchMethodException ex) {
                // getter 方法不存在，使用字段直接访问
            }
        }

        // 使用字段直接访问
        Field field = findField(clazz, fieldName);
        if (field != null) {
            field.setAccessible(true);
            return field.get(obj);
        }

        throw new NoSuchFieldException("Field '" + fieldName + "' not found in class " + clazz.getName());
    }

    /**
     * 获取索引访问的值（支持 List、Set、数组）
     *
     * @param obj   对象
     * @param index 索引
     * @return 索引处的值
     */
    private static Object getIndexedValue(Object obj, int index) {
        if (obj == null) {
            return null;
        }

        // 数组
        if (obj.getClass().isArray()) {
            int length = Array.getLength(obj);
            if (index < 0 || index >= length) {
                log.warn("Array index out of bounds: {} (length: {})", index, length);
                return null;
            }
            return Array.get(obj, index);
        }

        // List
        if (obj instanceof List) {
            List<?> list = (List<?>) obj;
            if (index < 0 || index >= list.size()) {
                log.warn("List index out of bounds: {} (size: {})", index, list.size());
                return null;
            }
            return list.get(index);
        }

        // Set（转为 List 后访问）
        if (obj instanceof Set) {
            Set<?> set = (Set<?>) obj;
            if (index < 0 || index >= set.size()) {
                log.warn("Set index out of bounds: {} (size: {})", index, set.size());
                return null;
            }
            return new ArrayList<>(set).get(index);
        }

        // Collection（转为 List 后访问）
        if (obj instanceof Collection) {
            Collection<?> collection = (Collection<?>) obj;
            if (index < 0 || index >= collection.size()) {
                log.warn("Collection index out of bounds: {} (size: {})", index, collection.size());
                return null;
            }
            return new ArrayList<>(collection).get(index);
        }

        log.warn("Object is not indexable: {}", obj.getClass().getName());
        return null;
    }

    /**
     * 获取 Map 中指定键的值
     *
     * @param obj Map 对象
     * @param key 键
     * @return 值
     */
    private static Object getMapValue(Object obj, String key) {
        if (obj == null) {
            return null;
        }

        if (obj instanceof Map) {
            return ((Map<?, ?>) obj).get(key);
        }

        log.warn("Object is not a Map: {}", obj.getClass().getName());
        return null;
    }

    /**
     * 拼接集合中所有元素的值
     *
     * @param obj         集合对象
     * @param separator   分隔符
     * @param maxJoinSize 最大拼接数量，0 表示不限制
     * @return 拼接后的字符串
     */
    private static Object joinCollectionValues(Object obj, String separator, int maxJoinSize) {
        if (obj == null) {
            return null;
        }

        Collection<Object> collection = null;

        // 数组转集合
        if (obj.getClass().isArray()) {
            int length = Array.getLength(obj);
            collection = new ArrayList<>(length);
            for (int i = 0; i < length; i++) {
                collection.add(Array.get(obj, i));
            }
        } else if (obj instanceof Collection) {
            collection = (Collection<Object>) obj;
        } else {
            log.warn("Object is not a collection or array: {}", obj.getClass().getName());
            return obj.toString();
        }

        // 限制数量
        if (maxJoinSize > 0 && collection.size() > maxJoinSize) {
            collection = collection.stream().limit(maxJoinSize).collect(Collectors.toList());
        }

        // 拼接
        return collection.stream()
                .filter(Objects::nonNull)
                .map(Object::toString)
                .collect(Collectors.joining(separator));
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
     * 检查字段是否有 NestedProperty 注解
     *
     * @param field 字段
     * @return 是否有注解
     */
    public static boolean hasNestedProperty(Field field) {
        return field.isAnnotationPresent(NestedProperty.class);
    }

    /**
     * 获取字段的 NestedProperty 注解
     *
     * @param field 字段
     * @return NestedProperty 注解，如果不存在则返回 null
     */
    public static NestedProperty getNestedProperty(Field field) {
        return field.getAnnotation(NestedProperty.class);
    }

    /**
     * 路径片段
     */
    private static class PathSegment {
        String fieldName;        // 字段名
        AccessorType accessorType; // 访问器类型
        Object accessor;         // 访问器参数（索引或键）

        PathSegment(String fieldName, AccessorType accessorType, Object accessor) {
            this.fieldName = fieldName;
            this.accessorType = accessorType;
            this.accessor = accessor;
        }
    }

    /**
     * 访问器类型
     */
    private enum AccessorType {
        FIELD,           // 字段访问：dept
        INDEX,           // 索引访问：depts[0]
        MAP_KEY,         // Map 键访问：properties[city]
        COLLECTION_ALL   // 集合全部访问：depts[*]
    }
}
