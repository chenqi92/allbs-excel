package cn.allbs.excel.template;

import cn.allbs.excel.constant.ExcelConstants;
import cn.idev.excel.FastExcel;
import cn.idev.excel.annotation.ExcelProperty;
import cn.idev.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * Excel Template Generator
 * <p>
 * Generates Excel import templates with sample data
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ExcelTemplateGenerator {

	private static final Random RANDOM = new Random();

	/**
	 * Generate template file with sample data
	 *
	 * @param outputPath    Output file path
	 * @param dataClass     Data class
	 * @param sampleRows    Number of sample rows (0 for header only)
	 * @param sheetName     Sheet name
	 * @param <T>           Data type
	 */
	public static <T> void generateTemplate(String outputPath, Class<T> dataClass, int sampleRows, String sheetName) {
		generateTemplate(new File(outputPath), dataClass, sampleRows, sheetName);
	}

	/**
	 * Generate template file with sample data
	 *
	 * @param outputFile    Output file
	 * @param dataClass     Data class
	 * @param sampleRows    Number of sample rows (0 for header only)
	 * @param sheetName     Sheet name
	 * @param <T>           Data type
	 */
	public static <T> void generateTemplate(File outputFile, Class<T> dataClass, int sampleRows, String sheetName) {
		List<T> sampleData = generateSampleData(dataClass, sampleRows);

		FastExcel.write(outputFile, dataClass)
			.registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
			.sheet(sheetName != null ? sheetName : ExcelConstants.DEFAULT_SHEET_NAME)
			.doWrite(sampleData);

		log.info("Template generated: {} (class: {}, sample rows: {})", outputFile.getAbsolutePath(),
				dataClass.getSimpleName(), sampleRows);
	}

	/**
	 * Generate template to output stream with sample data
	 *
	 * @param outputStream  Output stream
	 * @param dataClass     Data class
	 * @param sampleRows    Number of sample rows (0 for header only)
	 * @param sheetName     Sheet name
	 * @param <T>           Data type
	 */
	public static <T> void generateTemplate(OutputStream outputStream, Class<T> dataClass, int sampleRows,
			String sheetName) {
		List<T> sampleData = generateSampleData(dataClass, sampleRows);

		FastExcel.write(outputStream, dataClass)
			.registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
			.sheet(sheetName != null ? sheetName : ExcelConstants.DEFAULT_SHEET_NAME)
			.doWrite(sampleData);

		log.info("Template generated to stream (class: {}, sample rows: {})", dataClass.getSimpleName(), sampleRows);
	}

	/**
	 * Generate sample data
	 *
	 * @param dataClass     Data class
	 * @param rowCount      Number of rows
	 * @param <T>           Data type
	 * @return Sample data list
	 */
	public static <T> List<T> generateSampleData(Class<T> dataClass, int rowCount) {
		List<T> data = new ArrayList<>();

		if (rowCount <= 0) {
			return data;
		}

		try {
			// Find no-arg constructor
			Constructor<T> constructor = dataClass.getDeclaredConstructor();
			constructor.setAccessible(true);

			Field[] fields = dataClass.getDeclaredFields();

			for (int i = 0; i < rowCount; i++) {
				T instance = constructor.newInstance();

				for (Field field : fields) {
					ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
					if (excelProperty != null) {
						field.setAccessible(true);
						Object sampleValue = generateSampleValue(field.getType(), field.getName(), i);
						field.set(instance, sampleValue);
					}
				}

				data.add(instance);
			}

			log.debug("Generated {} sample rows for {}", rowCount, dataClass.getSimpleName());
		}
		catch (Exception e) {
			log.error("Failed to generate sample data for {}", dataClass.getSimpleName(), e);
		}

		return data;
	}

	/**
	 * Generate sample value for a field type
	 */
	private static Object generateSampleValue(Class<?> fieldType, String fieldName, int rowIndex) {
		String lowerFieldName = fieldName.toLowerCase();

		// String types
		if (fieldType == String.class) {
			if (lowerFieldName.contains("name")) {
				return "Sample Name " + (rowIndex + 1);
			}
			else if (lowerFieldName.contains("email")) {
				return "user" + (rowIndex + 1) + "@example.com";
			}
			else if (lowerFieldName.contains("phone")) {
				return "138" + String.format("%08d", RANDOM.nextInt(100000000));
			}
			else if (lowerFieldName.contains("address")) {
				return "Sample Address " + (rowIndex + 1);
			}
			else if (lowerFieldName.contains("id")) {
				return "ID" + String.format("%04d", rowIndex + 1);
			}
			else if (lowerFieldName.contains("code")) {
				return "CODE" + String.format("%04d", rowIndex + 1);
			}
			else {
				return "Sample Text " + (rowIndex + 1);
			}
		}

		// Integer types
		if (fieldType == Integer.class || fieldType == int.class) {
			if (lowerFieldName.contains("age")) {
				return 20 + RANDOM.nextInt(50);
			}
			else if (lowerFieldName.contains("quantity") || lowerFieldName.contains("count")) {
				return 1 + RANDOM.nextInt(100);
			}
			else if (lowerFieldName.contains("year")) {
				return 2020 + RANDOM.nextInt(5);
			}
			else {
				return (rowIndex + 1) * 10;
			}
		}

		// Long types
		if (fieldType == Long.class || fieldType == long.class) {
			return (long) (rowIndex + 1) * 1000;
		}

		// Double types
		if (fieldType == Double.class || fieldType == double.class) {
			if (lowerFieldName.contains("rate") || lowerFieldName.contains("percent")) {
				return 50.0 + RANDOM.nextDouble() * 50.0; // 50-100
			}
			else {
				return 100.0 + RANDOM.nextDouble() * 900.0; // 100-1000
			}
		}

		// Float types
		if (fieldType == Float.class || fieldType == float.class) {
			return 10.0f + RANDOM.nextFloat() * 90.0f; // 10-100
		}

		// BigDecimal types
		if (fieldType == BigDecimal.class) {
			if (lowerFieldName.contains("price") || lowerFieldName.contains("amount") || lowerFieldName.contains("salary")) {
				return new BigDecimal(1000 + RANDOM.nextInt(9000));
			}
			else {
				return new BigDecimal(100 + RANDOM.nextInt(900));
			}
		}

		// Boolean types
		if (fieldType == Boolean.class || fieldType == boolean.class) {
			return RANDOM.nextBoolean();
		}

		// Date types
		if (fieldType == Date.class) {
			return new Date();
		}

		// LocalDate types
		if (fieldType == LocalDate.class) {
			return LocalDate.now().minusDays(RANDOM.nextInt(365));
		}

		// LocalDateTime types
		if (fieldType == LocalDateTime.class) {
			return LocalDateTime.now().minusDays(RANDOM.nextInt(365));
		}

		// List types - generate empty list
		if (List.class.isAssignableFrom(fieldType)) {
			return new ArrayList<>();
		}

		// Default: null
		return null;
	}

}
