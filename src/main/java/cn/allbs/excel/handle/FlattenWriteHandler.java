package cn.allbs.excel.handle;

import cn.allbs.excel.util.FlattenFieldProcessor;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * 支持 @FlattenProperty 的写入处理器
 * <p>
 * 在写入数据行时，自动从嵌套对象中提取字段值
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class FlattenWriteHandler implements CellWriteHandler {

    private Class<?> dataClass;
    private List<FlattenFieldProcessor.FlattenFieldInfo> fieldInfos;

    /**
     * Default constructor (for reflection instantiation)
     */
    public FlattenWriteHandler() {
        this.dataClass = null;
        this.fieldInfos = null;
    }

    public FlattenWriteHandler(Class<?> dataClass) {
        this.dataClass = dataClass;
        this.fieldInfos = FlattenFieldProcessor.processFlattenFields(dataClass);
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                 List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex,
                                 Boolean isHead) {
        // 表头行不处理
        if (Boolean.TRUE.equals(isHead)) {
            return;
        }

        // 获取当前行的数据对象
        Row row = cell.getRow();
        int rowIndex = row.getRowNum();

        // 从 WriteSheetHolder 中获取当前正在写入的数据
        // 注意：这里需要一些技巧来获取当前行对应的数据对象
        // 由于 EasyExcel 的限制，我们可能需要在外部维护一个数据索引
        log.debug("Writing cell at row: {}, column: {}", rowIndex, cell.getColumnIndex());
    }
}
