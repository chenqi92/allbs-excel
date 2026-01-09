package cn.allbs.excel.annotation;

import cn.allbs.excel.head.HeadGenerator;
import cn.idev.excel.converters.Converter;
import cn.idev.excel.support.ExcelTypeEnum;
import cn.idev.excel.write.handler.WriteHandler;

import java.lang.annotation.*;

/**
 * 自定义注解功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
@Documented
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExportExcel {

    /**
     * 文件名称
     *
     * @return string
     */
    String name() default "";

    /**
     * 文件类型 （xlsx xls）
     *
     * @return string
     */
    ExcelTypeEnum suffix() default ExcelTypeEnum.XLSX;

    /**
     * 文件密码
     *
     * @return password
     */
    String password() default "";

    /**
     * sheet 名称，支持多个
     *
     * @return String[]
     */
    Sheet[] sheets() default @Sheet(sheetName = "sheet1");

    /**
     * 内存操作
     *
     * @return boolean
     */
    boolean inMemory() default false;

    /**
     * excel 模板
     *
     * @return String
     */
    String template() default "";

    /**
     * 包含字段
     *
     * @return String[]
     */
    String[] include() default {};

    /**
     * 排除字段
     *
     * @return String[]
     */
    String[] exclude() default {};

    /**
     * 拦截器，自定义样式等处理器
     *
     * @return WriteHandler[]
     */
    Class<? extends WriteHandler>[] writeHandler() default {};

    /**
     * 转换器
     *
     * @return Converter[]
     */
    Class<? extends Converter>[] converter() default {};

    /**
     * 自定义Excel头生成器
     *
     * @return HeadGenerator
     */
    Class<? extends HeadGenerator> headGenerator() default HeadGenerator.class;

    /**
     * excel 头信息国际化
     *
     * @return boolean
     */
    boolean i18nHeader() default false;

    /**
     * 填充模式
     *
     * @return boolean
     */
    boolean fill() default false;

    /**
     * 是否只导出有 @ExcelProperty 注解的字段
     * <p>
     * true: 只导出标注了 @ExcelProperty 的字段，忽略其他字段<br>
     * false: 导出所有字段（默认行为）
     * </p>
     * <p>
     * 等同于在实体类上添加 @ExcelIgnoreUnannotated 注解
     * </p>
     *
     * @return boolean
     */
    boolean onlyExcelProperty() default false;

    /**
     * 是否自动合并相同值的单元格
     * <p>
     * true: 自动合并标注了 @ExcelMerge 注解的字段的相同值单元格<br>
     * false: 不进行合并（默认行为）
     * </p>
     * <p>
     * 需要配合 @ExcelMerge 注解使用
     * </p>
     *
     * @return boolean
     */
    boolean autoMerge() default false;

    /**
     * Excel 水印配置
     * <p>
     * 如果配置了水印，会自动为所有 Sheet 添加水印
     * </p>
     *
     * @return ExcelWatermark
     */
    ExcelWatermark watermark() default @ExcelWatermark(text = "", enabled = false);

    /**
     * Excel 图表配置（单个图表）
     * <p>
     * 如果配置了图表，会自动在 Sheet 中生成图表
     * </p>
     *
     * @return ExcelChart
     */
    ExcelChart chart() default @ExcelChart(title = "", enabled = false);

    /**
     * Excel 多图表配置
     * <p>
     * 支持在同一个 Sheet 中生成多个图表
     * </p>
     *
     * @return ExcelChart[]
     */
    ExcelChart[] charts() default {};

}
