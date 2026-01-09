package cn.allbs.excel.handle;

import cn.allbs.excel.listener.ExportProgressListener;
import cn.idev.excel.write.handler.RowWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteSheetHolder;
import cn.idev.excel.write.metadata.holder.WriteTableHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;

/**
 * Excel 导出进度监听处理器
 * <p>
 * 实现 RowWriteHandler 接口，在每行写入后触发进度回调
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Slf4j
public class ProgressWriteHandler implements RowWriteHandler {

    /**
     * 进度监听器
     */
    private final ExportProgressListener listener;

    /**
     * 总行数
     */
    private final int totalRows;

    /**
     * 进度更新间隔
     */
    private final int interval;

    /**
     * Sheet 名称
     */
    private final String sheetName;

    /**
     * 当前已写入的数据行数（不包括表头）
     */
    private int currentRow = 0;

    /**
     * 是否已触发开始事件
     */
    private boolean started = false;

    /**
     * 构造函数
     *
     * @param listener  进度监听器
     * @param totalRows 总行数
     * @param interval  进度更新间隔
     * @param sheetName Sheet 名称
     */
    public ProgressWriteHandler(ExportProgressListener listener, int totalRows, int interval, String sheetName) {
        this.listener = listener;
        this.totalRows = totalRows;
        this.interval = interval;
        this.sheetName = sheetName;
    }

    @Override
    public void beforeRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                 Integer rowIndex, Integer relativeRowIndex, Boolean isHead) {
        // 不需要实现
    }

    @Override
    public void afterRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                Row row, Integer relativeRowIndex, Boolean isHead) {
        // 不需要实现
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                 Row row, Integer relativeRowIndex, Boolean isHead) {
        try {
            // 跳过表头
            if (isHead != null && isHead) {
                return;
            }

            // 第一次写入数据行时，触发开始事件
            if (!started) {
                started = true;
                listener.onStart(totalRows, sheetName);
                log.debug("Excel export started: totalRows={}, sheetName={}", totalRows, sheetName);
            }

            // 增加当前行数
            currentRow++;

            // 根据间隔触发进度回调
            if (interval > 0 && currentRow % interval == 0) {
                double percentage = totalRows > 0 ? (currentRow * 100.0 / totalRows) : 0;
                listener.onProgress(currentRow, totalRows, percentage, sheetName);
                log.debug("Excel export progress: {}/{} ({}%)", currentRow, totalRows, String.format("%.2f", percentage));
            }

            // 最后一行时触发完成事件
            if (currentRow >= totalRows) {
                listener.onComplete(totalRows, sheetName);
                log.debug("Excel export completed: totalRows={}, sheetName={}", totalRows, sheetName);
            }
        } catch (Exception e) {
            log.error("Error in progress callback", e);
            try {
                listener.onError(e, sheetName);
            } catch (Exception ex) {
                log.error("Error in error callback", ex);
            }
        }
    }
}

