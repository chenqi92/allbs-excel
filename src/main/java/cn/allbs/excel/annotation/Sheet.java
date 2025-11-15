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

}
