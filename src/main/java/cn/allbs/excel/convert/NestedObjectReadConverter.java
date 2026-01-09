package cn.allbs.excel.convert;

import cn.allbs.excel.annotation.NestedProperty;
import cn.idev.excel.converters.Converter;
import cn.idev.excel.metadata.GlobalConfiguration;
import cn.idev.excel.metadata.data.ReadCellData;
import cn.idev.excel.metadata.property.ExcelContentProperty;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * 嵌套对象导入转换器
 * <p>
 * 用于导入时将 Excel 列值设置回嵌套对象字段
 * </p>
 * <p>使用示例：</p>
 * <pre>
 * &#64;ExcelProperty(value = "部门名称", converter = NestedObjectReadConverter.class)
 * &#64;NestedProperty("name")
 * private Department department;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class NestedObjectReadConverter implements Converter<Object> {

	@Override
	public Object convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty,
			GlobalConfiguration globalConfiguration) {
		try {
			Field field = contentProperty.getField();
			NestedProperty annotation = field.getAnnotation(NestedProperty.class);

			if (annotation == null) {
				log.warn("Field {} does not have @NestedProperty annotation", field.getName());
				return null;
			}

			String path = annotation.value();
			String cellValue = cellData.getStringValue();

			if (cellValue == null || cellValue.isEmpty()) {
				return null;
			}

			// 创建嵌套对象并设置值
			Class<?> fieldType = field.getType();
			Object nestedObject = fieldType.getDeclaredConstructor().newInstance();

			// 解析路径并设置值
			setNestedValue(nestedObject, path, cellValue);

			return nestedObject;
		}
		catch (Exception e) {
			log.error("Failed to convert nested object", e);
			return null;
		}
	}

	/**
	 * 设置嵌套对象的值
	 *
	 * @param obj   对象
	 * @param path  路径（如 "name" 或 "leader.name"）
	 * @param value 值
	 */
	private void setNestedValue(Object obj, String path, String value) throws Exception {
		if (obj == null || path == null || path.isEmpty()) {
			return;
		}

		String[] parts = path.split("\\.", 2);
		String currentField = parts[0];

		// 简单属性，直接设置
		if (parts.length == 1) {
			setFieldValue(obj, currentField, value);
		}
		else {
			// 多层嵌套，递归处理
			Field field = findField(obj.getClass(), currentField);
			if (field == null) {
				log.warn("Field {} not found in {}", currentField, obj.getClass());
				return;
			}

			// 获取或创建嵌套对象
			Object nestedObj = getFieldValue(obj, field);
			if (nestedObj == null) {
				nestedObj = field.getType().getDeclaredConstructor().newInstance();
				setFieldValue(obj, field, nestedObj);
			}

			// 递归设置值
			setNestedValue(nestedObj, parts[1], value);
		}
	}

	/**
	 * 设置字段值
	 *
	 * @param obj       对象
	 * @param fieldName 字段名
	 * @param value     值（字符串）
	 */
	private void setFieldValue(Object obj, String fieldName, String value) throws Exception {
		Field field = findField(obj.getClass(), fieldName);
		if (field == null) {
			log.warn("Field {} not found", fieldName);
			return;
		}

		setFieldValue(obj, field, convertValue(value, field.getType()));
	}

	/**
	 * 设置字段值
	 *
	 * @param obj   对象
	 * @param field 字段
	 * @param value 值
	 */
	private void setFieldValue(Object obj, Field field, Object value) throws Exception {
		// 优先使用 setter 方法
		String setterName = "set" + capitalize(field.getName());
		try {
			Method setter = obj.getClass().getMethod(setterName, field.getType());
			setter.invoke(obj, value);
		}
		catch (NoSuchMethodException e) {
			// 直接设置字段
			field.setAccessible(true);
			field.set(obj, value);
		}
	}

	/**
	 * 获取字段值
	 *
	 * @param obj   对象
	 * @param field 字段
	 * @return 字段值
	 */
	private Object getFieldValue(Object obj, Field field) throws Exception {
		// 优先使用 getter 方法
		String getterName = "get" + capitalize(field.getName());
		try {
			Method getter = obj.getClass().getMethod(getterName);
			return getter.invoke(obj);
		}
		catch (NoSuchMethodException e) {
			// 尝试 is 开头的 getter
			try {
				String isGetterName = "is" + capitalize(field.getName());
				Method isGetter = obj.getClass().getMethod(isGetterName);
				return isGetter.invoke(obj);
			}
			catch (NoSuchMethodException ex) {
				// 直接访问字段
				field.setAccessible(true);
				return field.get(obj);
			}
		}
	}

	/**
	 * 查找字段
	 *
	 * @param clazz     类
	 * @param fieldName 字段名
	 * @return 字段
	 */
	private Field findField(Class<?> clazz, String fieldName) {
		try {
			return clazz.getDeclaredField(fieldName);
		}
		catch (NoSuchFieldException e) {
			// 查找父类
			if (clazz.getSuperclass() != null) {
				return findField(clazz.getSuperclass(), fieldName);
			}
			return null;
		}
	}

	/**
	 * 转换值类型
	 *
	 * @param value      字符串值
	 * @param targetType 目标类型
	 * @return 转换后的值
	 */
	private Object convertValue(String value, Class<?> targetType) {
		if (value == null) {
			return null;
		}

		if (targetType == String.class) {
			return value;
		}
		else if (targetType == Integer.class || targetType == int.class) {
			return Integer.parseInt(value);
		}
		else if (targetType == Long.class || targetType == long.class) {
			return Long.parseLong(value);
		}
		else if (targetType == Double.class || targetType == double.class) {
			return Double.parseDouble(value);
		}
		else if (targetType == Boolean.class || targetType == boolean.class) {
			return Boolean.parseBoolean(value);
		}
		else {
			// 默认返回字符串
			return value;
		}
	}

	/**
	 * 首字母大写
	 */
	private String capitalize(String str) {
		if (str == null || str.isEmpty()) {
			return str;
		}
		return str.substring(0, 1).toUpperCase() + str.substring(1);
	}

}
