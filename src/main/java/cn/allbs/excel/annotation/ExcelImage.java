package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 图片注解
 * <p>
 * 用于标记需要导出为图片的字段，支持单张图片或图片列表
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * // 单张图片
 * &#64;ExcelProperty("商品图片")
 * &#64;ExcelImage
 * private String imageUrl; // 可以是 URL 或本地路径
 *
 * // 单张图片（字节数组）
 * &#64;ExcelProperty("头像")
 * &#64;ExcelImage
 * private byte[] avatar;
 *
 * // 多张图片
 * &#64;ExcelProperty("商品图片集")
 * &#64;ExcelImage
 * private List&lt;String&gt; imageUrls;
 *
 * // 自定义图片大小
 * &#64;ExcelProperty("缩略图")
 * &#64;ExcelImage(width = 100, height = 100)
 * private String thumbnail;
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelImage {

	/**
	 * 图片宽度（像素）
	 * <p>
	 * 默认为 100 像素
	 * </p>
	 *
	 * @return 图片宽度
	 */
	int width() default 100;

	/**
	 * 图片高度（像素）
	 * <p>
	 * 默认为 100 像素
	 * </p>
	 *
	 * @return 图片高度
	 */
	int height() default 100;

	/**
	 * 图片类型
	 * <p>
	 * 支持的类型：
	 * </p>
	 * <ul>
	 *   <li>URL - 网络图片URL</li>
	 *   <li>FILE - 本地文件路径</li>
	 *   <li>BYTES - 字节数组</li>
	 *   <li>BASE64 - Base64编码的图片</li>
	 * </ul>
	 *
	 * @return 图片类型
	 */
	ImageType type() default ImageType.AUTO;

	/**
	 * 当图片加载失败时是否显示占位符
	 * <p>
	 * 如果为 true，当图片加载失败时会显示 "图片加载失败" 文本
	 * </p>
	 *
	 * @return 是否显示占位符
	 */
	boolean showPlaceholderOnError() default true;

	/**
	 * 图片在单元格中的位置
	 * <p>
	 * 当一个单元格有多张图片时，可以指定排列方式
	 * </p>
	 *
	 * @return 图片位置
	 */
	ImagePosition position() default ImagePosition.CENTER;

	/**
	 * 图片类型枚举
	 */
	enum ImageType {
		/**
		 * 自动检测（根据字段值判断）
		 */
		AUTO,
		/**
		 * 网络URL
		 */
		URL,
		/**
		 * 本地文件路径
		 */
		FILE,
		/**
		 * 字节数组
		 */
		BYTES,
		/**
		 * Base64编码
		 */
		BASE64
	}

	/**
	 * 图片位置枚举
	 */
	enum ImagePosition {
		/**
		 * 居中
		 */
		CENTER,
		/**
		 * 左对齐
		 */
		LEFT,
		/**
		 * 右对齐
		 */
		RIGHT,
		/**
		 * 顶部对齐
		 */
		TOP,
		/**
		 * 底部对齐
		 */
		BOTTOM
	}

}
