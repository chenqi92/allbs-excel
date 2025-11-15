package cn.allbs.excel.annotation;

import java.lang.annotation.*;

/**
 * Excel 导出进度回调注解
 * <p>
 * 用于标记需要进度回调的导出方法，配合 ExportProgressListener 使用
 * </p>
 * <p>
 * 使用示例：
 * <pre>
 * &#64;GetMapping("/export")
 * &#64;ExportExcel(name = "用户列表", sheets = &#64;Sheet(sheetName = "用户信息"))
 * &#64;ExportProgress(listener = MyProgressListener.class)
 * public List&lt;UserDTO&gt; export() {
 *     return userService.findAll();
 * }
 * </pre>
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Documented
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExportProgress {

    /**
     * 进度监听器类
     * <p>
     * 指定一个实现了 ExportProgressListener 接口的类，用于接收进度回调
     * </p>
     *
     * @return 监听器类
     */
    Class<? extends cn.allbs.excel.listener.ExportProgressListener> listener();

    /**
     * 进度更新间隔（行数）
     * <p>
     * 每导出多少行触发一次进度回调，默认为 100 行
     * </p>
     * <p>
     * 设置为 1 表示每行都触发回调（可能影响性能）<br>
     * 设置为 0 表示只在开始和结束时触发回调
     * </p>
     *
     * @return 进度更新间隔
     */
    int interval() default 100;

    /**
     * 是否启用进度回调
     * <p>
     * 可以通过此属性动态控制是否启用进度回调
     * </p>
     *
     * @return 是否启用
     */
    boolean enabled() default true;
}

