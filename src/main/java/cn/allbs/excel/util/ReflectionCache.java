package cn.allbs.excel.util;

import cn.allbs.excel.constant.ExcelConstants;
import cn.allbs.excel.exception.ExcelException;
import lombok.extern.slf4j.Slf4j;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Reflection Cache Utility
 * <p>
 * Caches reflection operations using MethodHandle for better performance (2-3x faster than regular reflection)
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ReflectionCache {

	private static final MethodHandles.Lookup LOOKUP = MethodHandles.lookup();

	/**
	 * Field getter cache: class#fieldName -> MethodHandle
	 */
	private static final Map<String, MethodHandle> GETTER_CACHE = new ConcurrentHashMap<>();

	/**
	 * Field setter cache: class#fieldName -> MethodHandle
	 */
	private static final Map<String, MethodHandle> SETTER_CACHE = new ConcurrentHashMap<>();

	/**
	 * Field cache: class#fieldName -> Field
	 */
	private static final Map<String, Field> FIELD_CACHE = new ConcurrentHashMap<>();

	/**
	 * Method cache: class#methodName#paramTypes -> Method
	 */
	private static final Map<String, Method> METHOD_CACHE = new ConcurrentHashMap<>();

	/**
	 * Cache statistics
	 */
	private static long hitCount = 0;

	private static long missCount = 0;

	/**
	 * Get field value using MethodHandle (much faster than Field.get())
	 *
	 * @param obj       Object instance
	 * @param fieldName Field name
	 * @return Field value
	 * @throws Throwable If error occurs
	 */
	public static Object getFieldValue(Object obj, String fieldName) throws Throwable {
		if (obj == null) {
			return null;
		}

		String key = getCacheKey(obj.getClass(), fieldName);
		MethodHandle getter = GETTER_CACHE.get(key);

		if (getter != null) {
			hitCount++;
			return getter.invoke(obj);
		}

		missCount++;
		getter = createGetter(obj.getClass(), fieldName);
		GETTER_CACHE.put(key, getter);

		return getter.invoke(obj);
	}

	/**
	 * Set field value using MethodHandle (much faster than Field.set())
	 *
	 * @param obj       Object instance
	 * @param fieldName Field name
	 * @param value     Value to set
	 * @throws Throwable If error occurs
	 */
	public static void setFieldValue(Object obj, String fieldName, Object value) throws Throwable {
		if (obj == null) {
			return;
		}

		String key = getCacheKey(obj.getClass(), fieldName);
		MethodHandle setter = SETTER_CACHE.get(key);

		if (setter != null) {
			hitCount++;
			setter.invoke(obj, value);
			return;
		}

		missCount++;
		setter = createSetter(obj.getClass(), fieldName);
		SETTER_CACHE.put(key, setter);

		setter.invoke(obj, value);
	}

	/**
	 * Get cached field
	 *
	 * @param clazz     Class
	 * @param fieldName Field name
	 * @return Field
	 */
	public static Field getCachedField(Class<?> clazz, String fieldName) {
		String key = getCacheKey(clazz, fieldName);
		Field field = FIELD_CACHE.get(key);

		if (field != null) {
			hitCount++;
			return field;
		}

		missCount++;
		field = findField(clazz, fieldName);
		if (field != null) {
			field.setAccessible(true);
			FIELD_CACHE.put(key, field);
		}

		return field;
	}

	/**
	 * Get cached method
	 *
	 * @param clazz      Class
	 * @param methodName Method name
	 * @param paramTypes Parameter types
	 * @return Method
	 */
	public static Method getCachedMethod(Class<?> clazz, String methodName, Class<?>... paramTypes) {
		String key = getMethodCacheKey(clazz, methodName, paramTypes);
		Method method = METHOD_CACHE.get(key);

		if (method != null) {
			hitCount++;
			return method;
		}

		missCount++;
		try {
			method = clazz.getMethod(methodName, paramTypes);
			method.setAccessible(true);
			METHOD_CACHE.put(key, method);
			return method;
		}
		catch (NoSuchMethodException e) {
			// Try declared method
			try {
				method = clazz.getDeclaredMethod(methodName, paramTypes);
				method.setAccessible(true);
				METHOD_CACHE.put(key, method);
				return method;
			}
			catch (NoSuchMethodException ex) {
				log.warn("Method not found: {}.{}({})", clazz.getName(), methodName,
						getParamTypeNames(paramTypes));
				return null;
			}
		}
	}

	/**
	 * Create getter MethodHandle
	 */
	private static MethodHandle createGetter(Class<?> clazz, String fieldName) {
		try {
			// Try to find getter method first (better performance)
			String getterName = "get" + capitalize(fieldName);
			try {
				Method getter = clazz.getMethod(getterName);
				return LOOKUP.unreflect(getter);
			}
			catch (NoSuchMethodException e) {
				// Try "is" prefix for boolean
				try {
					String isGetterName = "is" + capitalize(fieldName);
					Method isGetter = clazz.getMethod(isGetterName);
					return LOOKUP.unreflect(isGetter);
				}
				catch (NoSuchMethodException ex) {
					// Fall back to direct field access
					Field field = findField(clazz, fieldName);
					if (field == null) {
						throw new ExcelException(ExcelConstants.ERR_FIELD_NOT_FOUND,
								"Field not found: " + fieldName, fieldName);
					}
					field.setAccessible(true);
					return LOOKUP.unreflectGetter(field);
				}
			}
		}
		catch (IllegalAccessException e) {
			throw new ExcelException(ExcelConstants.ERR_FIELD_ACCESS_DENIED, "Cannot access field: " + fieldName, e,
					fieldName);
		}
	}

	/**
	 * Create setter MethodHandle
	 */
	private static MethodHandle createSetter(Class<?> clazz, String fieldName) {
		try {
			// Try to find setter method first
			Field field = findField(clazz, fieldName);
			if (field == null) {
				throw new ExcelException(ExcelConstants.ERR_FIELD_NOT_FOUND, "Field not found: " + fieldName,
						fieldName);
			}

			String setterName = "set" + capitalize(fieldName);
			try {
				Method setter = clazz.getMethod(setterName, field.getType());
				return LOOKUP.unreflect(setter);
			}
			catch (NoSuchMethodException e) {
				// Fall back to direct field access
				field.setAccessible(true);
				return LOOKUP.unreflectSetter(field);
			}
		}
		catch (IllegalAccessException e) {
			throw new ExcelException(ExcelConstants.ERR_FIELD_ACCESS_DENIED, "Cannot access field: " + fieldName, e,
					fieldName);
		}
	}

	/**
	 * Find field in class hierarchy
	 */
	private static Field findField(Class<?> clazz, String fieldName) {
		Class<?> current = clazz;
		while (current != null && current != Object.class) {
			try {
				return current.getDeclaredField(fieldName);
			}
			catch (NoSuchFieldException e) {
				current = current.getSuperclass();
			}
		}
		return null;
	}

	/**
	 * Capitalize first letter
	 */
	private static String capitalize(String str) {
		if (str == null || str.isEmpty()) {
			return str;
		}
		return Character.toUpperCase(str.charAt(0)) + str.substring(1);
	}

	/**
	 * Generate cache key
	 */
	private static String getCacheKey(Class<?> clazz, String fieldName) {
		return clazz.getName() + "#" + fieldName;
	}

	/**
	 * Generate method cache key
	 */
	private static String getMethodCacheKey(Class<?> clazz, String methodName, Class<?>... paramTypes) {
		StringBuilder sb = new StringBuilder();
		sb.append(clazz.getName()).append("#").append(methodName).append("#");
		for (Class<?> paramType : paramTypes) {
			sb.append(paramType.getName()).append(",");
		}
		return sb.toString();
	}

	/**
	 * Get parameter type names
	 */
	private static String getParamTypeNames(Class<?>... paramTypes) {
		if (paramTypes == null || paramTypes.length == 0) {
			return "";
		}
		StringBuilder sb = new StringBuilder();
		for (int i = 0; i < paramTypes.length; i++) {
			if (i > 0) {
				sb.append(", ");
			}
			sb.append(paramTypes[i].getSimpleName());
		}
		return sb.toString();
	}

	/**
	 * Get cache statistics
	 */
	public static String getCacheStats() {
		long total = hitCount + missCount;
		double hitRate = total > 0 ? (hitCount * 100.0 / total) : 0;
		return String.format(
				"ReflectionCache Stats - Getters: %d, Setters: %d, Fields: %d, Methods: %d, Hit Rate: %.2f%% (%d/%d)",
				GETTER_CACHE.size(), SETTER_CACHE.size(), FIELD_CACHE.size(), METHOD_CACHE.size(), hitRate, hitCount,
				total);
	}

	/**
	 * Clear all caches
	 */
	public static void clearCache() {
		GETTER_CACHE.clear();
		SETTER_CACHE.clear();
		FIELD_CACHE.clear();
		METHOD_CACHE.clear();
		hitCount = 0;
		missCount = 0;
		log.info("Reflection cache cleared");
	}

	/**
	 * Get cache size
	 */
	public static int getCacheSize() {
		return GETTER_CACHE.size() + SETTER_CACHE.size() + FIELD_CACHE.size() + METHOD_CACHE.size();
	}

}
