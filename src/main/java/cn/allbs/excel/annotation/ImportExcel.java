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

    /**
     * 图片读取模式
     * <p>
     * DRAWING: 读取Excel中的Drawing对象（真实图片）
     * BASE64: 读取单元格中的Base64文本
     * AUTO: 自动检测（默认使用DRAWING模式）
     * </p>
     *
     * @return 图片读取模式
     */
    ImageReadMode imageReadMode() default ImageReadMode.AUTO;

    /**
     * 图片读取模式枚举
     */
    enum ImageReadMode {
        /**
         * 读取Drawing对象（真实图片）
         */
        DRAWING,
        /**
         * 读取Base64文本
         */
        BASE64,
        /**
         * 自动检测（优先DRAWING）
         */
        AUTO
    }

}