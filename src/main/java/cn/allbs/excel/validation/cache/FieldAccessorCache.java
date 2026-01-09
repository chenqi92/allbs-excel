package cn.allbs.excel.validation.cache;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.reflect.Field;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 字段访问器缓存
 * <p>
 * 使用 MethodHandle 替代反射，提升字段访问性能
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class FieldAccessorCache {

    private static final Map<String, MethodHandle> ACCESSOR_CACHE = new ConcurrentHashMap<>();
    private static final Map<String, Field> FIELD_CACHE = new ConcurrentHashMap<>();
    private static final MethodHandles.Lookup LOOKUP = MethodHandles.lookup();

    private FieldAccessorCache() {
    }

    /**
     * 获取字段值
     *
     * @param obj       对象
     * @param fieldName 字段名
     * @return 字段值
     */
    public static Object getFieldValue(Object obj, String fieldName) {
        if (obj == null) {
            return null;
        }

        String key = obj.getClass().getName() + "#" + fieldName;
        MethodHandle handle = ACCESSOR_CACHE.computeIfAbsent(key, k -> createMethodHandle(obj.getClass(), fieldName));

        if (handle == null) {
            return null;
        }

        try {
            return handle.invoke(obj);
        } catch (Throwable e) {
            throw new RuntimeException("Failed to get field value: " + fieldName, e);
        }
    }

    /**
     * 获取字段对象
     *
     * @param clazz     类
     * @param fieldName 字段名
     * @return 字段对象
     */
    public static Field getField(Class<?> clazz, String fieldName) {
        String key = clazz.getName() + "#" + fieldName;
        return FIELD_CACHE.computeIfAbsent(key, k -> findField(clazz, fieldName));
    }

    /**
     * 创建 MethodHandle
     */
    private static MethodHandle createMethodHandle(Class<?> clazz, String fieldName) {
        Field field = getField(clazz, fieldName);
        if (field == null) {
            return null;
        }

        try {
            field.setAccessible(true);
            return LOOKUP.unreflectGetter(field);
        } catch (IllegalAccessException e) {
            throw new RuntimeException("Failed to create MethodHandle for field: " + fieldName, e);
        }
    }

    /**
     * 查找字段（包括父类）
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
     * 检查字段值是否为空
     * <p>
     * 支持 null、空字符串、空集合、空 Map、空数组
     * </p>
     *
     * @param value 字段值
     * @return 是否为空
     */
    public static boolean isEmpty(Object value) {
        if (value == null) {
            return true;
        }
        if (value instanceof String) {
            return ((String) value).trim().isEmpty();
        }
        if (value instanceof java.util.Collection) {
            return ((java.util.Collection<?>) value).isEmpty();
        }
        if (value instanceof java.util.Map) {
            return ((java.util.Map<?, ?>) value).isEmpty();
        }
        if (value.getClass().isArray()) {
            return java.lang.reflect.Array.getLength(value) == 0;
        }
        return false;
    }

    /**
     * 检查字段值是否非空
     *
     * @param value 字段值
     * @return 是否非空
     */
    public static boolean isNotEmpty(Object value) {
        return !isEmpty(value);
    }

    /**
     * 清除缓存（用于测试）
     */
    public static void clearCache() {
        ACCESSOR_CACHE.clear();
        FIELD_CACHE.clear();
    }
}
