package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelValidation;
import cn.allbs.excel.annotation.ValidationType;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel 数据验证写处理器
 * <p>
 * 为 Excel 列添加数据验证规则
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class ExcelValidationWriteHandler implements SheetWriteHandler {

	/**
	 * 数据类型
	 */
	private final Class<?> dataClass;

	/**
	 * 列索引 -> 验证信息
	 */
	private final Map<Integer, ValidationInfo> columnValidationMap = new HashMap<>();

	/**
	 * 验证的起始行（默认从第2行开始，第1行是表头）
	 */
	private final int firstRow;

	/**
	 * 验证的结束行（默认到第65535行）
	 */
	private final int lastRow;

	public ExcelValidationWriteHandler(Class<?> dataClass) {
		this(dataClass, 1, 65535);
	}

	public ExcelValidationWriteHandler(Class<?> dataClass, int firstRow, int lastRow) {
		this.dataClass = dataClass;
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		initValidations();
	}

	/**
	 * 初始化验证规则
	 */
	private void initValidations() {
		if (dataClass == null) {
			return;
		}

		Field[] fields = dataClass.getDeclaredFields();
		int columnIndex = 0;

		for (Field field : fields) {
			ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
			ExcelValidation validation = field.getAnnotation(ExcelValidation.class);

			if (excelProperty != null) {
				// 获取实际列索引
				int index = excelProperty.index() >= 0 ? excelProperty.index() : columnIndex;

				// 如果有验证注解且已启用，则记录验证信息
				if (validation != null && validation.enabled()) {
					ValidationInfo info = new ValidationInfo();
					info.field = field;
					info.validation = validation;
					columnValidationMap.put(index, info);
					log.debug("Registered validation for column {}: {}", index, field.getName());
				}

				// 正确更新 columnIndex：取当前 index + 1 和 columnIndex 的最大值
				// 这样可以正确处理非连续的显式 index (如 0, 2, 5, 10)
				columnIndex = Math.max(columnIndex, index + 1);
			}
		}

		log.info("Initialized validations for {} columns", columnValidationMap.size());
	}

	@Override
	public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
		Sheet sheet = writeSheetHolder.getSheet();
		DataValidationHelper helper = sheet.getDataValidationHelper();

		for (Map.Entry<Integer, ValidationInfo> entry : columnValidationMap.entrySet()) {
			int columnIndex = entry.getKey();
			ValidationInfo info = entry.getValue();
			ExcelValidation validation = info.validation;

			try {
				// 创建验证约束
				DataValidationConstraint constraint = createConstraint(helper, validation);
				if (constraint == null) {
					continue;
				}

				// 创建验证区域
				CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, columnIndex,
						columnIndex);

				// 创建数据验证
				DataValidation dataValidation = helper.createValidation(constraint, regions);

				// 设置错误提示
				if (validation.showErrorBox()) {
					dataValidation.setShowErrorBox(true);
					dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
					dataValidation.createErrorBox(validation.errorTitle(), validation.errorMessage());
				}

				// 设置输入提示
				if (validation.showPromptBox() && !validation.promptMessage().isEmpty()) {
					dataValidation.setShowPromptBox(true);
					dataValidation.createPromptBox(validation.promptTitle(), validation.promptMessage());
				}

				// 应用验证
				sheet.addValidationData(dataValidation);

				log.debug("Added validation for column {}: {}", columnIndex, validation.type());
			}
			catch (Exception e) {
				log.error("Failed to add validation for column {}", columnIndex, e);
			}
		}
	}

	/**
	 * 创建验证约束
	 */
	private DataValidationConstraint createConstraint(DataValidationHelper helper, ExcelValidation validation) {
		ValidationType type = validation.type();

		switch (type) {
			case LIST:
				// 下拉列表
				if (validation.options().length == 0) {
					log.warn("LIST validation requires options");
					return null;
				}
				return helper.createExplicitListConstraint(validation.options());

			case NUMBER_RANGE:
			case INTEGER:
			case DECIMAL:
				// 数值范围
				int operatorType = DataValidationConstraint.OperatorType.BETWEEN;
				int validationType = type == ValidationType.INTEGER ? DataValidationConstraint.ValidationType.INTEGER
						: DataValidationConstraint.ValidationType.DECIMAL;

				if (validation.min() != Double.MIN_VALUE && validation.max() != Double.MAX_VALUE) {
					// 范围验证
					return helper.createNumericConstraint(validationType, operatorType,
							String.valueOf(validation.min()), String.valueOf(validation.max()));
				}
				else if (validation.min() != Double.MIN_VALUE) {
					// 最小值验证
					return helper.createNumericConstraint(validationType,
							DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, String.valueOf(validation.min()),
							null);
				}
				else if (validation.max() != Double.MAX_VALUE) {
					// 最大值验证
					return helper.createNumericConstraint(validationType,
							DataValidationConstraint.OperatorType.LESS_OR_EQUAL, String.valueOf(validation.max()),
							null);
				}
				return null;

			case DATE:
			case TIME:
				// 日期/时间验证
				return helper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "1900-01-01",
						"2099-12-31", validation.dateFormat());

			case TEXT_LENGTH:
				// 文本长度验证
				if (validation.minLength() > 0 && validation.maxLength() < Integer.MAX_VALUE) {
					return helper.createTextLengthConstraint(DataValidationConstraint.OperatorType.BETWEEN,
							String.valueOf(validation.minLength()), String.valueOf(validation.maxLength()));
				}
				else if (validation.minLength() > 0) {
					return helper.createTextLengthConstraint(
							DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
							String.valueOf(validation.minLength()), null);
				}
				else if (validation.maxLength() < Integer.MAX_VALUE) {
					return helper.createTextLengthConstraint(DataValidationConstraint.OperatorType.LESS_OR_EQUAL,
							String.valueOf(validation.maxLength()), null);
				}
				return null;

			case FORMULA:
				// 自定义公式
				if (validation.formula().isEmpty()) {
					log.warn("FORMULA validation requires formula");
					return null;
				}
				return helper.createCustomConstraint(validation.formula());

			case ANY:
				// 任意值（仅用于显示提示）
				return helper.createCustomConstraint("TRUE");

			default:
				log.warn("Unsupported validation type: {}", type);
				return null;
		}
	}

	/**
	 * 验证信息
	 */
	private static class ValidationInfo {

		Field field;

		ExcelValidation validation;

	}

}
