package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 数据验证注解
 * <p>
 * 用于在导出的 Excel 中添加数据验证规则
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 下拉列表
 * &#64;ExcelProperty("性别")
 * &#64;ExcelValidation(type = ValidationType.LIST, options = {"男", "女"})
 * private String gender;
 *
 * // 数值范围
 * &#64;ExcelProperty("年龄")
 * &#64;ExcelValidation(type = ValidationType.NUMBER_RANGE, min = 18, max = 65, errorMessage = "年龄必须在18-65之间")
 * private Integer age;
 *
 * // 日期范围
 * &#64;ExcelProperty("入职日期")
 * &#64;ExcelValidation(type = ValidationType.DATE, dateFormat = "yyyy-MM-dd", promptMessage = "请输入日期")
 * private LocalDate hireDate;
 *
 * // 文本长度
 * &#64;ExcelProperty("姓名")
 * &#64;ExcelValidation(type = ValidationType.TEXT_LENGTH, minLength = 2, maxLength = 10)
 * private String name;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelValidation {

	/**
	 * 验证类型
	 *
	 * @return 验证类型
	 */
	ValidationType type();

	/**
	 * 下拉列表选项（type = LIST 时使用）
	 * <p>
	 * 例如：{"男", "女"}、{"是", "否"}
	 * </p>
	 *
	 * @return 选项数组
	 */
	String[] options() default {};

	/**
	 * 最小值（type = NUMBER_RANGE、INTEGER、DECIMAL 时使用）
	 *
	 * @return 最小值
	 */
	double min() default Double.MIN_VALUE;

	/**
	 * 最大值（type = NUMBER_RANGE、INTEGER、DECIMAL 时使用）
	 *
	 * @return 最大值
	 */
	double max() default Double.MAX_VALUE;

	/**
	 * 最小长度（type = TEXT_LENGTH 时使用）
	 *
	 * @return 最小长度
	 */
	int minLength() default 0;

	/**
	 * 最大长度（type = TEXT_LENGTH 时使用）
	 *
	 * @return 最大长度
	 */
	int maxLength() default Integer.MAX_VALUE;

	/**
	 * 日期格式（type = DATE、TIME 时使用）
	 *
	 * @return 日期格式
	 */
	String dateFormat() default "yyyy-MM-dd";

	/**
	 * 自定义公式（type = FORMULA 时使用）
	 * <p>
	 * 例如："{current_cell}>=18" 表示当前单元格值必须大于等于18
	 * </p>
	 *
	 * @return 公式
	 */
	String formula() default "";

	/**
	 * 错误提示消息
	 * <p>
	 * 当用户输入无效数据时显示的错误信息
	 * </p>
	 *
	 * @return 错误消息
	 */
	String errorMessage() default "输入的数据无效";

	/**
	 * 错误提示标题
	 *
	 * @return 错误标题
	 */
	String errorTitle() default "数据验证错误";

	/**
	 * 输入提示消息
	 * <p>
	 * 当用户选中单元格时显示的提示信息
	 * </p>
	 *
	 * @return 提示消息
	 */
	String promptMessage() default "";

	/**
	 * 输入提示标题
	 *
	 * @return 提示标题
	 */
	String promptTitle() default "输入提示";

	/**
	 * 是否显示错误警告
	 * <p>
	 * true: 显示错误对话框（阻止输入）<br>
	 * false: 仅显示警告（允许输入）
	 * </p>
	 *
	 * @return 是否显示错误警告
	 */
	boolean showErrorBox() default true;

	/**
	 * 是否显示输入提示
	 *
	 * @return 是否显示输入提示
	 */
	boolean showPromptBox() default false;

	/**
	 * 是否启用
	 *
	 * @return 是否启用
	 */
	boolean enabled() default true;

}
