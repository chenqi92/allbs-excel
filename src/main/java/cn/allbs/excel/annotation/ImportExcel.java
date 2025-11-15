package cn.allbs.excel.annotation;


import cn.allbs.excel.handle.DefaultAnalysisEventListener;
import cn.allbs.excel.handle.ListAnalysisEventListener;

import java.lang.annotation.*;

/**
 * excel 导入
 *
 * @author ChenQi
 */
@Documented
@Target({ElementType.PARAMETER})
@Retention(RetentionPolicy.RUNTIME)
public @interface ImportExcel {

    /**
     * 前端上传字段名称 file
     */
    String fileName() default "file";

    /**
     * 读取的监听器类
     *
     * @return readListener
     */
    Class<? extends ListAnalysisEventListener<?>> readListener() default DefaultAnalysisEventListener.class;

    /**
     * 是否跳过空行
     *
     * @return 默认跳过
     */
    boolean ignoreEmptyRow() default false;

}