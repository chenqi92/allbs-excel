package cn.allbs.excel.annotation;


import cn.allbs.excel.head.HeadGenerator;

import java.lang.annotation.*;

/**
 * @author ChenQi
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Sheet {

    int sheetNo() default -1;

    /**
     * sheet name
     */
    String sheetName();

    /**
     * 包含字段
     */
    String[] includes() default {};

    /**
     * 排除字段
     */
    String[] excludes() default {};

    /**
     * 头生成器
     */
    Class<? extends HeadGenerator> headGenerateClass() default HeadGenerator.class;

    /**
     * 数据类型（用于空数据时生成表头）
     * 当返回空List时，如果需要导出带表头的Excel，必须指定此属性
     */
    Class<?> clazz() default Void.class;

    /**
     * 是否只导出有 @ExcelProperty 注解的字段
     * <p>
     * true: 只导出标注了 @ExcelProperty 的字段，忽略其他字段<br>
     * false: 导出所有字段（默认行为）
     * </p>
     * <p>
     * 等同于在实体类上添加 @ExcelIgnoreUnannotated 注解
     * </p>
     * <p>
     * 优先级：Sheet 级别配置 > ExportExcel 级别配置
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
     * <p>
     * 优先级：Sheet 级别配置 > ExportExcel 级别配置
     * </p>
     *
     * @return boolean
     */
    boolean autoMerge() default false;

    /**
     * 数据验证起始行（从0开始，0表示第一行数据）
     * <p>
     * 用于自定义 ExcelValidation 的验证范围
     * </p>
     * <p>
     * -1 表示使用默认值（从第一行数据开始）
     * </p>
     *
     * @return int
     */
    int validationStartRow() default -1;

    /**
     * 数据验证结束行（从0开始）
     * <p>
     * 用于自定义 ExcelValidation 的验证范围
     * </p>
     * <p>
     * -1 表示使用默认值（到最后一行）
     * </p>
     *
     * @return int
     */
    int validationEndRow() default -1;

}
