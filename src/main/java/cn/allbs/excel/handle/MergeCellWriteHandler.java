package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelMerge;
import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 单元格合并处理器
 * <p>
 * 自动合并相同值的单元格，支持依赖关系
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
@Slf4j
public class MergeCellWriteHandler implements RowWriteHandler, SheetWriteHandler, WorkbookWriteHandler {

    /**
     * 数据类型
     */
    private Class<?> dataClass;

    /**
     * 需要合并的列索引 -> 字段信息
     */
    private final Map<Integer, MergeFieldInfo> mergeColumnMap = new HashMap<>();

    /**
     * 所有数据行的缓存（用于最后统一合并）
     */
    private final List<RowData> rowDataList = new ArrayList<>();

    /**
     * 表头行数
     */
    private int headRowNumber = 1;

    /**
     * 是否已经执行过合并
     */
    private boolean merged = false;

    /**
     * 保存 Sheet 对象的引用
     */
    private Sheet sheet;

    /**
     * Default constructor (for reflection instantiation)
     */
    public MergeCellWriteHandler() {
        this.dataClass = null;
    }

    public MergeCellWriteHandler(Class<?> dataClass) {
        this.dataClass = dataClass;
        initMergeColumns();
    }

    /**
     * 初始化需要合并的列
     */
    private void initMergeColumns() {
        if (dataClass == null) {
            return;
        }

        Field[] fields = dataClass.getDeclaredFields();
        for (Field field : fields) {
            ExcelMerge excelMerge = field.getAnnotation(ExcelMerge.class);
            if (excelMerge != null && excelMerge.enabled()) {
                // 获取字段的列索引
                com.alibaba.excel.annotation.ExcelProperty excelProperty =
                        field.getAnnotation(com.alibaba.excel.annotation.ExcelProperty.class);
                if (excelProperty != null) {
                    int columnIndex = excelProperty.index();
                    MergeFieldInfo fieldInfo = new MergeFieldInfo(
                            field.getName(),
                            columnIndex,
                            excelMerge.dependOn()
                    );
                    mergeColumnMap.put(columnIndex, fieldInfo);
                }
            }
        }
    }

    // ========== WorkbookWriteHandler 接口方法 ==========

    @Override
    public void beforeWorkbookCreate() {
        // 不需要实现
    }

    @Override
    public void afterWorkbookCreate(WriteWorkbookHolder writeWorkbookHolder) {
        // 不需要实现
    }

    @Override
    public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
        // 在工作簿完成前执行合并操作
        if (!merged && sheet != null && !mergeColumnMap.isEmpty() && !rowDataList.isEmpty()) {
            log.debug("开始执行单元格合并，共 {} 列需要合并，{} 行数据", mergeColumnMap.size(), rowDataList.size());
            doMerge(sheet);
            merged = true;
        }
    }

    // ========== SheetWriteHandler 接口方法 ==========

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        // 不需要实现
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        // 保存 Sheet 对象的引用，用于后续合并操作
        this.sheet = writeSheetHolder.getSheet();
    }

    // ========== RowWriteHandler 接口方法 ==========

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
        // 记录表头行数
        if (isHead != null && isHead) {
            headRowNumber = Math.max(headRowNumber, row.getRowNum() + 1);
            return;
        }

        // 如果没有需要合并的列，直接返回
        if (mergeColumnMap.isEmpty()) {
            return;
        }

        // 收集数据行信息
        int rowIndex = row.getRowNum();

        // 确保 rowDataList 有足够的容量
        while (rowDataList.size() <= rowIndex - headRowNumber) {
            rowDataList.add(new RowData(rowIndex));
        }

        RowData rowData = rowDataList.get(rowIndex - headRowNumber);

        // 遍历所有需要合并的列，收集数据
        for (Integer columnIndex : mergeColumnMap.keySet()) {
            Cell cell = row.getCell(columnIndex);
            String cellValue = getCellValue(cell);
            rowData.setCellValue(columnIndex, cellValue);
        }
    }



    /**
     * 执行合并操作
     * <p>
     * 这个方法需要在所有数据写入完成后手动调用
     * </p>
     *
     * @param sheet 工作表
     */
    public void doMerge(Sheet sheet) {
        if (mergeColumnMap.isEmpty() || rowDataList.isEmpty()) {
            return;
        }

        // 按列索引排序，确保依赖关系正确处理
        List<Integer> sortedColumns = new ArrayList<>(mergeColumnMap.keySet());
        Collections.sort(sortedColumns);

        for (Integer columnIndex : sortedColumns) {
            MergeFieldInfo fieldInfo = mergeColumnMap.get(columnIndex);
            mergeSameValueCells(sheet, columnIndex, fieldInfo);
        }
    }

    /**
     * 合并相同值的单元格
     *
     * @param sheet       工作表
     * @param columnIndex 列索引
     * @param fieldInfo   字段信息
     */
    private void mergeSameValueCells(Sheet sheet, int columnIndex, MergeFieldInfo fieldInfo) {
        int startRow = headRowNumber;
        int endRow = headRowNumber + rowDataList.size() - 1;

        int mergeStartRow = startRow;
        String previousValue = null;

        for (int i = startRow; i <= endRow + 1; i++) {
            String currentValue = null;
            boolean shouldMerge = false;

            if (i <= endRow) {
                RowData rowData = rowDataList.get(i - headRowNumber);
                currentValue = rowData.getCellValue(columnIndex);

                // 检查依赖字段是否相同
                if (fieldInfo.dependOn.length > 0) {
                    boolean dependOnSame = true;
                    if (i > startRow) {
                        RowData previousRowData = rowDataList.get(i - headRowNumber - 1);
                        for (String dependField : fieldInfo.dependOn) {
                            Integer dependColumnIndex = findColumnIndexByFieldName(dependField);
                            if (dependColumnIndex != null) {
                                String currentDependValue = rowData.getCellValue(dependColumnIndex);
                                String previousDependValue = previousRowData.getCellValue(dependColumnIndex);
                                if (!Objects.equals(currentDependValue, previousDependValue)) {
                                    dependOnSame = false;
                                    break;
                                }
                            }
                        }
                    }
                    if (!dependOnSame) {
                        shouldMerge = true;
                    }
                }
            }

            // 判断是否需要合并
            if (i == endRow + 1 || shouldMerge || !Objects.equals(currentValue, previousValue)) {
                // 合并前面的单元格
                if (i - mergeStartRow > 1) {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(
                            mergeStartRow, i - 1, columnIndex, columnIndex
                    );
                    sheet.addMergedRegion(cellRangeAddress);
                }
                mergeStartRow = i;
            }

            previousValue = currentValue;
        }
    }

    /**
     * 根据字段名查找列索引
     *
     * @param fieldName 字段名
     * @return 列索引
     */
    private Integer findColumnIndexByFieldName(String fieldName) {
        for (Map.Entry<Integer, MergeFieldInfo> entry : mergeColumnMap.entrySet()) {
            if (entry.getValue().fieldName.equals(fieldName)) {
                return entry.getKey();
            }
        }

        // 如果在合并列中没找到，尝试从数据类中查找
        try {
            Field field = dataClass.getDeclaredField(fieldName);
            com.alibaba.excel.annotation.ExcelProperty excelProperty =
                    field.getAnnotation(com.alibaba.excel.annotation.ExcelProperty.class);
            if (excelProperty != null) {
                return excelProperty.index();
            }
        } catch (NoSuchFieldException e) {
            log.warn("字段 {} 不存在", fieldName);
        }

        return null;
    }

    /**
     * 获取单元格的值
     *
     * @param cell 单元格
     * @return 单元格值
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return null;
        }
    }

    /**
     * 合并字段信息
     */
    private static class MergeFieldInfo {
        /**
         * 字段名
         */
        private final String fieldName;

        /**
         * 列索引
         */
        private final int columnIndex;

        /**
         * 依赖的字段名
         */
        private final String[] dependOn;

        public MergeFieldInfo(String fieldName, int columnIndex, String[] dependOn) {
            this.fieldName = fieldName;
            this.columnIndex = columnIndex;
            this.dependOn = dependOn;
        }
    }

    /**
     * 行数据
     */
    private static class RowData {
        /**
         * 行索引
         */
        private final int rowIndex;

        /**
         * 列索引 -> 单元格值
         */
        private final Map<Integer, String> cellValues = new HashMap<>();

        public RowData(int rowIndex) {
            this.rowIndex = rowIndex;
        }

        public void setCellValue(int columnIndex, String value) {
            cellValues.put(columnIndex, value);
        }

        public String getCellValue(int columnIndex) {
            return cellValues.get(columnIndex);
        }
    }
}

