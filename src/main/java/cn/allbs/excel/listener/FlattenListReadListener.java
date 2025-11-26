package cn.allbs.excel.listener;

import cn.allbs.excel.annotation.FlattenList;
import cn.allbs.excel.annotation.FlattenProperty;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;

/**
 * FlattenList 导入聚合监听器
 * <p>
 * 将多行 Excel 数据聚合回包含 List 的对象
 * </p>
 * <p>使用示例：</p>
 * <pre>
 * FlattenListReadListener&lt;FlattenListOrderDTO&gt; listener =
 *     new FlattenListReadListener&lt;&gt;(FlattenListOrderDTO.class);
 * EasyExcel.read(file, listener).sheet().doRead();
 * List&lt;FlattenListOrderDTO&gt; result = listener.getResult();
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class FlattenListReadListener<T> extends AnalysisEventListener<Map<Integer, String>> {

	private final Class<T> dataClass;

	/**
	 * 聚合后的结果列表
	 */
	@Getter
	private final List<T> result = new ArrayList<>();

	/**
	 * 当前正在聚合的对象
	 */
	private T currentObject;

	/**
	 * 上一行的普通字段值（用于判断是否是同一组）
	 */
	private Map<String, Object> previousNormalFields;

	/**
	 * 元数据
	 */
	private AggregateMetadata metadata;

	/**
	 * 表头映射（列索引 -> 列名）
	 */
	private Map<Integer, String> headMap;

	public FlattenListReadListener(Class<T> dataClass) {
		this.dataClass = dataClass;
		this.metadata = analyzeClass(dataClass);
	}

	@Override
	public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
		this.headMap = headMap;
		log.info("Excel headers: {}", headMap);
	}

	@Override
	public void invoke(Map<Integer, String> data, AnalysisContext context) {
		try {
			// 提取普通字段值
			Map<String, Object> normalFields = extractNormalFields(data);

			// 判断是否需要创建新对象
			if (currentObject == null || !isSameGroup(normalFields, previousNormalFields)) {
				// 保存上一个对象
				if (currentObject != null) {
					result.add(currentObject);
				}

				// 创建新对象
				currentObject = dataClass.getDeclaredConstructor().newInstance();

				// 设置普通字段
				setNormalFields(currentObject, normalFields);

				// 设置 @FlattenProperty 嵌套对象字段
				setFlattenPropertyFields(currentObject, data);

				// 初始化 List 字段
				initializeListFields(currentObject);

				previousNormalFields = normalFields;
			}

			// 添加 List 元素
			addListElements(currentObject, data);
		}
		catch (Exception e) {
			log.error("Failed to process row", e);
		}
	}

	@Override
	public void doAfterAllAnalysed(AnalysisContext context) {
		// 保存最后一个对象
		if (currentObject != null) {
			result.add(currentObject);
		}
		log.info("Import completed, total {} records", result.size());
	}

	/**
	 * 分析类结构
	 */
	private AggregateMetadata analyzeClass(Class<T> clazz) {
		AggregateMetadata meta = new AggregateMetadata();

		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			if (field.isAnnotationPresent(FlattenList.class)) {
				FlattenList annotation = field.getAnnotation(FlattenList.class);
				ListFieldMeta listMeta = new ListFieldMeta();
				listMeta.field = field;
				listMeta.annotation = annotation;

				// 获取泛型类型
				Type genericType = field.getGenericType();
				if (genericType instanceof ParameterizedType) {
					ParameterizedType paramType = (ParameterizedType) genericType;
					Type[] actualTypes = paramType.getActualTypeArguments();
					if (actualTypes.length > 0 && actualTypes[0] instanceof Class) {
						listMeta.elementType = (Class<?>) actualTypes[0];
						listMeta.elementFields = scanElementFields((Class<?>) actualTypes[0], annotation.prefix(), annotation.suffix());
					}
				}

				meta.listFields.add(listMeta);
			}
			else if (field.isAnnotationPresent(FlattenProperty.class)) {
				// 处理 @FlattenProperty 注解的嵌套对象
				FlattenProperty annotation = field.getAnnotation(FlattenProperty.class);
				FlattenPropertyMeta flattenMeta = new FlattenPropertyMeta();
				flattenMeta.field = field;
				flattenMeta.annotation = annotation;
				flattenMeta.nestedFields = scanElementFields(field.getType(), annotation.prefix(), annotation.suffix());
				meta.flattenPropertyFields.add(flattenMeta);
			}
			else if (field.isAnnotationPresent(ExcelProperty.class)) {
				NormalFieldMeta normalMeta = new NormalFieldMeta();
				normalMeta.field = field;
				normalMeta.excelProperty = field.getAnnotation(ExcelProperty.class);
				meta.normalFields.add(normalMeta);
			}
		}

		return meta;
	}

	/**
	 * 扫描元素字段
	 */
	private List<ElementFieldMeta> scanElementFields(Class<?> elementClass, String prefix, String suffix) {
		List<ElementFieldMeta> result = new ArrayList<>();
		Field[] fields = elementClass.getDeclaredFields();

		for (Field field : fields) {
			if (field.isAnnotationPresent(ExcelProperty.class)) {
				ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
				String[] headNames = excelProperty.value();
				String headName = headNames.length > 0 ? headNames[0] : field.getName();

				// 应用前缀和后缀
				String finalHeadName = prefix + headName + suffix;

				ElementFieldMeta elementMeta = new ElementFieldMeta();
				elementMeta.field = field;
				elementMeta.excelHeaderName = finalHeadName;

				result.add(elementMeta);
			}
		}

		return result;
	}

	/**
	 * 提取普通字段值
	 */
	private Map<String, Object> extractNormalFields(Map<Integer, String> data) {
		Map<String, Object> result = new HashMap<>();

		for (NormalFieldMeta normalMeta : metadata.normalFields) {
			String[] headNames = normalMeta.excelProperty.value();
			String headName = headNames.length > 0 ? headNames[0] : normalMeta.field.getName();

			// 查找对应的列索引
			Integer columnIndex = findColumnIndex(headName);
			if (columnIndex != null && data.containsKey(columnIndex)) {
				result.put(normalMeta.field.getName(), data.get(columnIndex));
			}
		}

		return result;
	}

	/**
	 * 判断是否是同一组
	 */
	private boolean isSameGroup(Map<String, Object> current, Map<String, Object> previous) {
		if (previous == null) {
			return false;
		}

		for (String key : current.keySet()) {
			if (!Objects.equals(current.get(key), previous.get(key))) {
				return false;
			}
		}

		return true;
	}

	/**
	 * 设置普通字段
	 */
	private void setNormalFields(T obj, Map<String, Object> normalFields) throws Exception {
		for (NormalFieldMeta normalMeta : metadata.normalFields) {
			Object value = normalFields.get(normalMeta.field.getName());
			if (value != null) {
				setFieldValue(obj, normalMeta.field, value);
			}
		}
	}

	/**
	 * 设置 @FlattenProperty 嵌套对象字段
	 */
	private void setFlattenPropertyFields(T obj, Map<Integer, String> data) throws Exception {
		for (FlattenPropertyMeta flattenMeta : metadata.flattenPropertyFields) {
			// 创建嵌套对象实例
			Object nestedObj = flattenMeta.field.getType().getDeclaredConstructor().newInstance();
			boolean hasData = false;

			// 设置嵌套对象的字段值
			for (ElementFieldMeta elementMeta : flattenMeta.nestedFields) {
				Integer columnIndex = findColumnIndex(elementMeta.excelHeaderName);
				if (columnIndex != null && data.containsKey(columnIndex)) {
					String value = data.get(columnIndex);
					if (value != null && !value.isEmpty()) {
						setFieldValue(nestedObj, elementMeta.field, value);
						hasData = true;
					}
				}
			}

			// 只设置有数据的嵌套对象
			if (hasData) {
				setFieldValue(obj, flattenMeta.field, nestedObj);
			}
		}
	}

	/**
	 * 初始化 List 字段
	 */
	private void initializeListFields(T obj) throws Exception {
		for (ListFieldMeta listMeta : metadata.listFields) {
			List<Object> list = new ArrayList<>();
			setFieldValue(obj, listMeta.field, list);
		}
	}

	/**
	 * 添加 List 元素
	 */
	@SuppressWarnings("unchecked")
	private void addListElements(T obj, Map<Integer, String> data) throws Exception {
		for (ListFieldMeta listMeta : metadata.listFields) {
			// 获取 List
			List<Object> list = (List<Object>) getFieldValue(obj, listMeta.field);

			// 创建元素对象
			Object element = listMeta.elementType.getDeclaredConstructor().newInstance();

			// 设置元素字段
			boolean hasData = false;
			for (ElementFieldMeta elementMeta : listMeta.elementFields) {
				Integer columnIndex = findColumnIndex(elementMeta.excelHeaderName);
				if (columnIndex != null && data.containsKey(columnIndex)) {
					String value = data.get(columnIndex);
					if (value != null && !value.isEmpty()) {
						setFieldValue(element, elementMeta.field, value);
						hasData = true;
					}
				}
			}

			// 只添加有数据的元素
			if (hasData) {
				list.add(element);
			}
		}
	}

	/**
	 * 查找列索引
	 */
	private Integer findColumnIndex(String headerName) {
		for (Map.Entry<Integer, String> entry : headMap.entrySet()) {
			if (entry.getValue().equals(headerName)) {
				return entry.getKey();
			}
		}
		return null;
	}

	/**
	 * 设置字段值
	 */
	private void setFieldValue(Object obj, Field field, Object value) throws Exception {
		String setterName = "set" + capitalize(field.getName());

		// 确定最终要设置的值
		Object finalValue;
		if (value == null) {
			finalValue = null;
		} else if (field.getType().isAssignableFrom(value.getClass())) {
			// 值已经是目标类型或其子类，无需转换
			finalValue = value;
		} else if (value instanceof String) {
			// 只有 String 类型才需要转换
			finalValue = convertValue((String) value, field.getType());
		} else {
			// 其他情况尝试直接使用
			finalValue = value;
		}

		try {
			Method setter = obj.getClass().getMethod(setterName, field.getType());
			setter.invoke(obj, finalValue);
		}
		catch (NoSuchMethodException e) {
			field.setAccessible(true);
			field.set(obj, finalValue);
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
		}
		catch (NoSuchMethodException e) {
			field.setAccessible(true);
			return field.get(obj);
		}
	}

	/**
	 * 常用日期时间格式
	 */
	private static final DateTimeFormatter[] DATE_TIME_FORMATTERS = {
			DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
			DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"),
			DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss"),
			DateTimeFormatter.ISO_LOCAL_DATE_TIME
	};

	private static final DateTimeFormatter[] DATE_FORMATTERS = {
			DateTimeFormatter.ofPattern("yyyy-MM-dd"),
			DateTimeFormatter.ofPattern("yyyy/MM/dd"),
			DateTimeFormatter.ISO_LOCAL_DATE
	};

	/**
	 * 转换值类型
	 */
	private Object convertValue(String value, Class<?> targetType) {
		if (value == null || value.isEmpty()) {
			return null;
		}

		try {
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
			else if (targetType == Float.class || targetType == float.class) {
				return Float.parseFloat(value);
			}
			else if (targetType == Boolean.class || targetType == boolean.class) {
				return Boolean.parseBoolean(value);
			}
			else if (targetType == BigDecimal.class) {
				return new BigDecimal(value);
			}
			else if (targetType == Short.class || targetType == short.class) {
				return Short.parseShort(value);
			}
			else if (targetType == Byte.class || targetType == byte.class) {
				return Byte.parseByte(value);
			}
			else if (targetType == LocalDateTime.class) {
				return parseLocalDateTime(value);
			}
			else if (targetType == LocalDate.class) {
				return parseLocalDate(value);
			}
			else {
				// 不支持的类型，返回 null
				log.debug("Unsupported field type: {}, value: {}", targetType.getName(), value);
				return null;
			}
		} catch (Exception e) {
			log.warn("Failed to convert value '{}' to type {}: {}", value, targetType.getName(), e.getMessage());
			return null;
		}
	}

	/**
	 * 解析 LocalDateTime
	 */
	private LocalDateTime parseLocalDateTime(String value) {
		for (DateTimeFormatter formatter : DATE_TIME_FORMATTERS) {
			try {
				return LocalDateTime.parse(value, formatter);
			} catch (DateTimeParseException ignored) {
				// 尝试下一个格式
			}
		}
		log.warn("Failed to parse LocalDateTime: {}", value);
		return null;
	}

	/**
	 * 解析 LocalDate
	 */
	private LocalDate parseLocalDate(String value) {
		for (DateTimeFormatter formatter : DATE_FORMATTERS) {
			try {
				return LocalDate.parse(value, formatter);
			} catch (DateTimeParseException ignored) {
				// 尝试下一个格式
			}
		}
		log.warn("Failed to parse LocalDate: {}", value);
		return null;
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

	// ==================== 内部类 ====================

	private static class AggregateMetadata {

		List<NormalFieldMeta> normalFields = new ArrayList<>();

		List<ListFieldMeta> listFields = new ArrayList<>();

		List<FlattenPropertyMeta> flattenPropertyFields = new ArrayList<>();

	}

	private static class NormalFieldMeta {

		Field field;

		ExcelProperty excelProperty;

	}

	private static class ListFieldMeta {

		Field field;

		FlattenList annotation;

		Class<?> elementType;

		List<ElementFieldMeta> elementFields;

	}

	private static class ElementFieldMeta {

		Field field;

		String excelHeaderName;

	}

	private static class FlattenPropertyMeta {

		Field field;

		FlattenProperty annotation;

		List<ElementFieldMeta> nestedFields;

	}

}
