package cn.allbs.excel.listener;

/**
 * Excel 导出进度监听器
 * <p>
 * 用于监听 Excel 导出过程中的进度变化，支持实时获取导出进度
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
public interface ExportProgressListener {

    /**
     * 导出开始
     *
     * @param totalRows 总行数（不包括表头）
     * @param sheetName Sheet 名称
     */
    void onStart(int totalRows, String sheetName);

    /**
     * 导出进度更新
     *
     * @param currentRow 当前已导出的行数
     * @param totalRows  总行数
     * @param percentage 进度百分比（0-100）
     * @param sheetName  Sheet 名称
     */
    void onProgress(int currentRow, int totalRows, double percentage, String sheetName);

    /**
     * 导出完成
     *
     * @param totalRows 总行数
     * @param sheetName Sheet 名称
     */
    void onComplete(int totalRows, String sheetName);

    /**
     * 导出失败
     *
     * @param exception 异常信息
     * @param sheetName Sheet 名称
     */
    void onError(Exception exception, String sheetName);
}

