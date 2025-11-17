package cn.allbs.excel.annotation;

/**
 * Excel 数据验证类型
 *
 * @author ChenQi
 * @since 2025-11-17
 */
public enum ValidationType {

	/**
	 * 下拉列表
	 */
	LIST,

	/**
	 * 数值范围
	 */
	NUMBER_RANGE,

	/**
	 * 整数
	 */
	INTEGER,

	/**
	 * 小数
	 */
	DECIMAL,

	/**
	 * 日期
	 */
	DATE,

	/**
	 * 时间
	 */
	TIME,

	/**
	 * 文本长度
	 */
	TEXT_LENGTH,

	/**
	 * 自定义公式
	 */
	FORMULA,

	/**
	 * 任意值（用于提示信息）
	 */
	ANY

}
