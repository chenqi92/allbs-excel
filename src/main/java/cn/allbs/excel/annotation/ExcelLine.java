package cn.allbs.excel.annotation;

import java.lang.annotation.*;

/**
 * 验证行号问题
 *
 * @author ChenQi
 */
@Documented
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelLine {

}
