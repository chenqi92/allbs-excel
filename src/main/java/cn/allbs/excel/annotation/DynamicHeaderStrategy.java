package cn.allbs.excel.annotation;

/**
 * 动态表头生成策略
 *
 * @author ChenQi
 * @since 2025-11-17
 */
public enum DynamicHeaderStrategy {

	/**
	 * 从数据中提取表头
	 * <p>
	 * 遍历所有数据行，收集 Map/Collection 中的所有键作为表头
	 * </p>
	 */
	FROM_DATA,

	/**
	 * 从配置中读取表头
	 * <p>
	 * 从注解的 headers 属性中读取预定义的表头列表
	 * </p>
	 */
	FROM_CONFIG,

	/**
	 * 混合模式
	 * <p>
	 * 先使用配置的表头，再补充数据中存在但配置中没有的表头
	 * </p>
	 */
	MIXED

}
