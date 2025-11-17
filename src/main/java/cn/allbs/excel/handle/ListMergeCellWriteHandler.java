package cn.allbs.excel.handle;

import cn.allbs.excel.util.ListEntityExpander;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * List 展开合并单元格处理器
 * <p>
 * 用于合并 @FlattenList 展开后的非 List 字段单元格
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class ListMergeCellWriteHandler implements SheetWriteHandler {

    private final List<ListEntityExpander.MergeRegion> mergeRegions;

    public ListMergeCellWriteHandler(List<ListEntityExpander.MergeRegion> mergeRegions) {
        this.mergeRegions = mergeRegions;
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Sheet sheet = writeSheetHolder.getSheet();

        if (mergeRegions == null || mergeRegions.isEmpty()) {
            return;
        }

        log.info("Merging {} regions", mergeRegions.size());

        for (ListEntityExpander.MergeRegion region : mergeRegions) {
            try {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(
                        region.getFirstRow(),
                        region.getLastRow(),
                        region.getFirstColumn(),
                        region.getLastColumn()
                );
                sheet.addMergedRegion(cellRangeAddress);

                log.debug("Merged region: ({},{}) to ({},{})",
                        region.getFirstRow(), region.getFirstColumn(),
                        region.getLastRow(), region.getLastColumn());
            } catch (Exception e) {
                log.error("Failed to merge region: {}", region, e);
            }
        }
    }
}
