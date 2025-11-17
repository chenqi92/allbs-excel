package cn.allbs.excel.listener;

import cn.allbs.excel.annotation.FlattenList;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
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
						listMeta.elementFields = scanElementFields((Class<?>) actualTypes[0], annotation);
					}
				}

				meta.listFields.add(listMeta);
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
	private List<ElementFieldMeta> scanElementFields(Class<?> elementClass, FlattenList annotation) {
		List<ElementFieldMeta> result = new ArrayList<>();
		Field[] fields = elementClass.getDeclaredFields();

		for (Field field : fields) {
			if (field.isAnnotationPresent(ExcelProperty.class)) {
				ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
				String[] headNames = excelProperty.value();
				String headName = headNames.length > 0 ? headNames[0] : field.getName();

				// 应用前缀和后缀
				String finalHeadName = annotation.prefix() + headName + annotation.suffix();

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
		try {
			// 类型转换
			Object convertedValue = convertValue(value.toString(), field.getType());
			Method setter = obj.getClass().getMethod(setterName, field.getType());
			setter.invoke(obj, convertedValue);
		}
		catch (NoSuchMethodException e) {
			field.setAccessible(true);
			Object convertedValue = convertValue(value.toString(), field.getType());
			field.set(obj, convertedValue);
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
	 * 转换值类型
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

	// ==================== 内部类 ====================

	private static class AggregateMetadata {

		List<NormalFieldMeta> normalFields = new ArrayList<>();

		List<ListFieldMeta> listFields = new ArrayList<>();

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

}
