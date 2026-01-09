package cn.allbs.excel.util;

import cn.allbs.excel.annotation.DynamicHeaders;
import cn.allbs.excel.annotation.DynamicHeaderStrategy;
import cn.idev.excel.annotation.ExcelProperty;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 动态表头处理器
 * <p>
 * 处理带有 @DynamicHeaders 注解的字段，生成动态表头并展开数据
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class DynamicHeaderProcessor {

	/**
	 * 展开数据列表
	 *
	 * @param dataList 原始数据列表
	 * @param <T>      数据类型
	 * @return 展开后的数据列表（使用 Map 表示）
	 */
	public static <T> List<Map<String, Object>> expandData(List<T> dataList) {
		if (dataList == null || dataList.isEmpty()) {
			return Collections.emptyList();
		}

		Class<?> dataClass = dataList.get(0).getClass();
		DynamicHeaderMetadata metadata = analyzeClass(dataClass, dataList);

		List<Map<String, Object>> result = new ArrayList<>();

		for (T data : dataList) {
			Map<String, Object> row = expandSingleObject(data, metadata);
			result.add(row);
		}

		return result;
	}

	/**
	 * 分析类结构，提取动态表头字段和普通字段信息
	 *
	 * @param clazz    数据类
	 * @param dataList 数据列表（用于从数据中提取动态表头）
	 * @return 动态表头元数据
	 */
	public static DynamicHeaderMetadata analyzeClass(Class<?> clazz, List<?> dataList) {
		DynamicHeaderMetadata metadata = new DynamicHeaderMetadata();
		metadata.setDataClass(clazz);

		Field[] fields = clazz.getDeclaredFields();

		for (Field field : fields) {
			if (field.isAnnotationPresent(DynamicHeaders.class)) {
				// 动态表头字段
				DynamicHeaders annotation = field.getAnnotation(DynamicHeaders.class);
				if (!annotation.enabled()) {
					continue;
				}

				DynamicHeaderFieldInfo fieldInfo = new DynamicHeaderFieldInfo();
				fieldInfo.setField(field);
				fieldInfo.setAnnotation(annotation);

				// 生成动态表头
				List<String> dynamicHeaders = generateDynamicHeaders(annotation, dataList, field);
				fieldInfo.setDynamicHeaders(dynamicHeaders);

				metadata.getDynamicHeaderFields().add(fieldInfo);
			}
			else if (field.isAnnotationPresent(ExcelProperty.class)) {
				// 普通字段
				NormalFieldInfo normalFieldInfo = new NormalFieldInfo();
				normalFieldInfo.setField(field);
				normalFieldInfo.setExcelProperty(field.getAnnotation(ExcelProperty.class));
				metadata.getNormalFields().add(normalFieldInfo);
			}
		}

		return metadata;
	}

	/**
	 * 生成动态表头
	 *
	 * @param annotation 注解
	 * @param dataList   数据列表
	 * @param field      字段
	 * @return 动态表头列表
	 */
	private static List<String> generateDynamicHeaders(DynamicHeaders annotation, List<?> dataList, Field field) {
		Set<String> headerSet = new LinkedHashSet<>();

		DynamicHeaderStrategy strategy = annotation.strategy();

		// 1. FROM_CONFIG 或 MIXED：添加预定义表头
		if (strategy == DynamicHeaderStrategy.FROM_CONFIG || strategy == DynamicHeaderStrategy.MIXED) {
			String[] predefinedHeaders = annotation.headers();
			for (String header : predefinedHeaders) {
				String finalHeader = annotation.headerPrefix() + header + annotation.headerSuffix();
				headerSet.add(finalHeader);
			}
		}

		// 2. FROM_DATA 或 MIXED：从数据中提取表头
		if (strategy == DynamicHeaderStrategy.FROM_DATA || strategy == DynamicHeaderStrategy.MIXED) {
			for (Object data : dataList) {
				try {
					Object fieldValue = getFieldValue(data, field);
					if (fieldValue instanceof Map) {
						Map<?, ?> map = (Map<?, ?>) fieldValue;
						for (Object key : map.keySet()) {
							if (key != null) {
								String header = annotation.headerPrefix() + key.toString()
										+ annotation.headerSuffix();
								headerSet.add(header);
							}
						}
					}
				}
				catch (Exception e) {
					log.error("Failed to extract dynamic headers from field: {}", field.getName(), e);
				}
			}
		}

		// 3. 转换为 List
		List<String> headers = new ArrayList<>(headerSet);

		// 4. 排序
		if (annotation.order() == DynamicHeaders.SortOrder.ASC) {
			Collections.sort(headers);
		}
		else if (annotation.order() == DynamicHeaders.SortOrder.DESC) {
			headers.sort(Collections.reverseOrder());
		}

		// 5. 限制列数
		if (annotation.maxColumns() > 0 && headers.size() > annotation.maxColumns()) {
			headers = headers.subList(0, annotation.maxColumns());
		}

		return headers;
	}

	/**
	 * 展开单个对象
	 *
	 * @param obj      对象
	 * @param metadata 元数据
	 * @return 展开后的行
	 */
	private static Map<String, Object> expandSingleObject(Object obj, DynamicHeaderMetadata metadata) {
		Map<String, Object> row = new LinkedHashMap<>();

		try {
			// 1. 添加普通字段
			for (NormalFieldInfo normalField : metadata.getNormalFields()) {
				Object value = getFieldValue(obj, normalField.getField());
				String[] headNames = normalField.getExcelProperty().value();
				String key = headNames.length > 0 ? headNames[0] : normalField.getField().getName();
				row.put(key, value);
			}

			// 2. 添加动态表头字段
			for (DynamicHeaderFieldInfo dynamicField : metadata.getDynamicHeaderFields()) {
				Object fieldValue = getFieldValue(obj, dynamicField.getField());
				Map<String, Object> dynamicData = extractDynamicData(fieldValue, dynamicField);

				// 按照动态表头顺序添加值
				for (String header : dynamicField.getDynamicHeaders()) {
					// 移除前缀和后缀得到原始键
					String originalKey = removePrexAndSuffix(header, dynamicField.getAnnotation().headerPrefix(),
							dynamicField.getAnnotation().headerSuffix());
					Object value = dynamicData.get(originalKey);
					row.put(header, value);
				}
			}
		}
		catch (Exception e) {
			log.error("Failed to expand object", e);
		}

		return row;
	}

	/**
	 * 提取动态数据
	 *
	 * @param fieldValue   字段值
	 * @param dynamicField 动态字段信息
	 * @return 动态数据 Map
	 */
	@SuppressWarnings("unchecked")
	private static Map<String, Object> extractDynamicData(Object fieldValue, DynamicHeaderFieldInfo dynamicField) {
		if (fieldValue == null) {
			return Collections.emptyMap();
		}

		if (fieldValue instanceof Map) {
			Map<String, Object> result = new LinkedHashMap<>();
			((Map<?, ?>) fieldValue).forEach((k, v) -> {
				if (k != null) {
					result.put(k.toString(), v);
				}
			});
			return result;
		}

		return Collections.emptyMap();
	}

	/**
	 * 移除前缀和后缀
	 *
	 * @param header 表头
	 * @param prefix 前缀
	 * @param suffix 后缀
	 * @return 原始键
	 */
	private static String removePrexAndSuffix(String header, String prefix, String suffix) {
		String result = header;
		if (!prefix.isEmpty() && result.startsWith(prefix)) {
			result = result.substring(prefix.length());
		}
		if (!suffix.isEmpty() && result.endsWith(suffix)) {
			result = result.substring(0, result.length() - suffix.length());
		}
		return result;
	}

	/**
	 * 获取字段值
	 *
	 * @param obj   对象
	 * @param field 字段
	 * @return 字段值
	 */
	private static Object getFieldValue(Object obj, Field field) {
		try {
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
		catch (Exception e) {
			log.error("Failed to get field value: {}", field.getName(), e);
			return null;
		}
	}

	/**
	 * 首字母大写
	 */
	private static String capitalize(String str) {
		if (str == null || str.isEmpty()) {
			return str;
		}
		return str.substring(0, 1).toUpperCase() + str.substring(1);
	}

	/**
	 * 生成表头列名列表
	 *
	 * @param metadata 元数据
	 * @return 表头列表
	 */
	public static List<String> generateHeaders(DynamicHeaderMetadata metadata) {
		List<String> headers = new ArrayList<>();

		// 添加普通字段表头
		for (NormalFieldInfo normalField : metadata.getNormalFields()) {
			String[] headNames = normalField.getExcelProperty().value();
			String headName = headNames.length > 0 ? headNames[0] : normalField.getField().getName();
			headers.add(headName);
		}

		// 添加动态字段表头
		for (DynamicHeaderFieldInfo dynamicField : metadata.getDynamicHeaderFields()) {
			headers.addAll(dynamicField.getDynamicHeaders());
		}

		return headers;
	}

	// ==================== 数据类 ====================

	/**
	 * 动态表头元数据
	 */
	@Data
	public static class DynamicHeaderMetadata {

		private Class<?> dataClass;

		private List<NormalFieldInfo> normalFields = new ArrayList<>();

		private List<DynamicHeaderFieldInfo> dynamicHeaderFields = new ArrayList<>();

	}

	/**
	 * 普通字段信息
	 */
	@Data
	public static class NormalFieldInfo {

		private Field field;

		private ExcelProperty excelProperty;

	}

	/**
	 * 动态表头字段信息
	 */
	@Data
	public static class DynamicHeaderFieldInfo {

		private Field field;

		private DynamicHeaders annotation;

		private List<String> dynamicHeaders;

	}

}
