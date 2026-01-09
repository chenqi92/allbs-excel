package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExcelLine;
import cn.idev.excel.write.handler.RowWriteHandler;
import cn.idev.excel.write.metadata.holder.WriteSheetHolder;
import cn.idev.excel.write.metadata.holder.WriteTableHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * Excel Line Number Write Handler
 * <p>
 * Automatically fills row numbers for fields annotated with @ExcelLine
 * Row numbers start from 1 for the first data row (excluding header)
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-23
 */
@Slf4j
public class ExcelLineWriteHandler implements RowWriteHandler {

    /**
     * Data class
     */
    private final Class<?> dataClass;

    /**
     * Column index for @ExcelLine field
     */
    private Integer lineColumnIndex = null;

    /**
     * Row number counter (starts from 1)
     */
    private int currentRowNumber = 1;

    /**
     * Constructor
     *
     * @param dataClass Data class
     */
    public ExcelLineWriteHandler(Class<?> dataClass) {
        this.dataClass = dataClass;
        initLineColumn();
    }

    /**
     * Initialize line column index
     */
    private void initLineColumn() {
        if (dataClass == null) {
            return;
        }

        Field[] fields = dataClass.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (field.isAnnotationPresent(ExcelLine.class)) {
                // Find the column index by checking @ExcelProperty annotation
                cn.idev.excel.annotation.ExcelProperty excelProperty =
                    field.getAnnotation(cn.idev.excel.annotation.ExcelProperty.class);

                if (excelProperty != null) {
                    lineColumnIndex = excelProperty.index() >= 0 ? excelProperty.index() : i;
                    log.info("Found @ExcelLine field '{}' at column index {}", field.getName(), lineColumnIndex);
                    break;
                }
            }
        }
    }

    @Override
    public void beforeRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                Integer rowIndex, Integer relativeRowIndex, Boolean isHead) {
        // No action needed
    }

    @Override
    public void afterRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                               Row row, Integer relativeRowIndex, Boolean isHead) {
        // No action needed
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder,
                                Row row, Integer relativeRowIndex, Boolean isHead) {
        // Skip header row
        if (isHead != null && isHead) {
            return;
        }

        // Skip if no @ExcelLine field found
        if (lineColumnIndex == null) {
            return;
        }

        // Fill row number
        Cell cell = row.getCell(lineColumnIndex);
        if (cell == null) {
            cell = row.createCell(lineColumnIndex);
        }

        // Set row number (starts from 1)
        cell.setCellValue(currentRowNumber);
        log.info("Set row number {} at row index {}, column index {}",
                 currentRowNumber, row.getRowNum(), lineColumnIndex);

        // Increment counter
        currentRowNumber++;
    }
}
