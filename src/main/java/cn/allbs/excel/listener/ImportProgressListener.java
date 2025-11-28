package cn.allbs.excel.listener;

/**
 * Excel 导入进度监听器
 * <p>
 * 用于监听 Excel 导入过程中的进度变化，支持实时获取导入进度
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-28
 */
public interface ImportProgressListener {

    /**
     * 导入开始
     *
     * @param totalRows 总行数（预估，可能不准确）
     * @param sheetName Sheet 名称
     */
    void onStart(int totalRows, String sheetName);

    /**
     * 导入进度更新
     *
     * @param currentRow 当前已导入的行数
     * @param totalRows  总行数（如果已知）
     * @param percentage 进度百分比（0-100）
     * @param sheetName  Sheet 名称
     */
    void onProgress(int currentRow, int totalRows, double percentage, String sheetName);

    /**
     * 导入完成
     *
     * @param totalRows 实际导入的总行数
     * @param sheetName Sheet 名称
     */
    void onComplete(int totalRows, String sheetName);

    /**
     * 导入失败
     *
     * @param exception 异常信息
     * @param sheetName Sheet 名称
     */
    void onError(Exception exception, String sheetName);
}
