package cn.allbs.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.data.ReadCellData;
import lombok.extern.slf4j.Slf4j;

import java.util.Map;

/**
 * Excel 导入进度监听处理器
 * <p>
 * 包装原有的 ReadListener，在读取过程中触发进度回调
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-28
 */
@Slf4j
public class ProgressReadListener<T> extends AnalysisEventListener<T> {

    /**
     * 进度监听器
     */
    private final ImportProgressListener progressListener;

    /**
     * 原始的读取监听器
     */
    private final AnalysisEventListener<T> delegate;

    /**
     * 进度更新间隔
     */
    private final int interval;

    /**
     * Sheet 名称
     */
    private String sheetName = "Sheet1";

    /**
     * 当前已读取的数据行数（不包括表头）
     */
    private int currentRow = 0;

    /**
     * 总行数（预估）
     */
    private int totalRows = 0;

    /**
     * 是否已触发开始事件
     */
    private boolean started = false;

    /**
     * 上次更新进度时的百分比（用于避免重复发送相同百分比）
     */
    private double lastPercentage = -1;

    /**
     * 构造函数
     *
     * @param progressListener 进度监听器
     * @param delegate         原始的读取监听器
     * @param interval         进度更新间隔
     */
    public ProgressReadListener(ImportProgressListener progressListener,
                                 AnalysisEventListener<T> delegate,
                                 int interval) {
        this.progressListener = progressListener;
        this.delegate = delegate;
        this.interval = interval;
    }

    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        // 获取 Sheet 名称
        if (context.readSheetHolder() != null) {
            sheetName = context.readSheetHolder().getSheetName();
        }

        // 尝试获取总行数（EasyExcel 可能无法准确获取）
        try {
            if (context.readSheetHolder() != null && context.readSheetHolder().getApproximateTotalRowNumber() != null) {
                totalRows = context.readSheetHolder().getApproximateTotalRowNumber() - 1; // 减去表头
            }
        } catch (Exception e) {
            log.debug("无法获取总行数，将使用动态计算", e);
        }

        // 触发开始事件
        if (!started) {
            started = true;
            progressListener.onStart(totalRows, sheetName);
            log.debug("Excel import started: totalRows={}, sheetName={}", totalRows, sheetName);
        }

        // 调用原始监听器
        if (delegate != null) {
            delegate.invokeHead(headMap, context);
        }
    }

    @Override
    public void invoke(T data, AnalysisContext context) {
        try {
            // 如果还没触发开始事件（没有表头的情况）
            if (!started) {
                if (context.readSheetHolder() != null) {
                    sheetName = context.readSheetHolder().getSheetName();
                }
                started = true;
                progressListener.onStart(totalRows, sheetName);
                log.debug("Excel import started: totalRows={}, sheetName={}", totalRows, sheetName);
            }

            // 增加当前行数
            currentRow++;

            // 计算当前百分比
            double percentage = totalRows > 0 ? (currentRow * 100.0 / totalRows) : 0;

            // 根据百分比变化或间隔触发进度回调
            // 策略：百分比变化超过0.5%或达到interval行数时更新
            boolean shouldUpdate = false;
            if (totalRows > 0) {
                // 基于百分比变化的更新（每0.5%更新一次）
                shouldUpdate = (percentage - lastPercentage) >= 0.5;
            } else {
                // 如果无法获取总行数，使用传统的行数间隔
                shouldUpdate = (interval > 0 && currentRow % interval == 0);
            }

            if (shouldUpdate) {
                lastPercentage = percentage;
                progressListener.onProgress(currentRow, totalRows, percentage, sheetName);
                log.debug("Excel import progress: {}/{} ({}%)", currentRow, totalRows, String.format("%.2f", percentage));
            }

            // 调用原始监听器
            if (delegate != null) {
                delegate.invoke(data, context);
            }
        } catch (Exception e) {
            log.error("Error in progress callback during invoke", e);
            progressListener.onError(e, sheetName);
            throw e;
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        try {
            // 触发完成事件
            double percentage = 100.0;
            if (currentRow > 0) {
                progressListener.onProgress(currentRow, currentRow, percentage, sheetName);
            }
            progressListener.onComplete(currentRow, sheetName);
            log.debug("Excel import completed: totalRows={}, sheetName={}", currentRow, sheetName);

            // 调用原始监听器
            if (delegate != null) {
                delegate.doAfterAllAnalysed(context);
            }
        } catch (Exception e) {
            log.error("Error in progress callback during completion", e);
            progressListener.onError(e, sheetName);
            throw e;
        }
    }

    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        log.error("Error during Excel import", exception);
        progressListener.onError(exception, sheetName);

        // 调用原始监听器
        if (delegate != null) {
            delegate.onException(exception, context);
        } else {
            throw exception;
        }
    }

    @Override
    public boolean hasNext(AnalysisContext context) {
        if (delegate != null) {
            return delegate.hasNext(context);
        }
        return super.hasNext(context);
    }
}
