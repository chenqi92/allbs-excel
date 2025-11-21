package cn.allbs.excel.handle;

import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.annotation.FlattenList;
import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.enhance.WriterBuilderEnhancer;
import cn.allbs.excel.kit.ExcelException;
import cn.allbs.excel.util.ListEntityExpander;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.ObjectProvider;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * FlattenList 处理器
 * <p>
 * 自动检测实体类中的 @FlattenList 注解，并使用 ListEntityExpander 处理 List 字段展开
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-21
 */
public class FlattenListWriteHandler extends AbstractSheetWriteHandler {

    public FlattenListWriteHandler(ExcelConfigProperties configProperties,
                                   ObjectProvider<List<Converter<?>>> converterProvider,
                                   WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * 支持条件：
     * 1. 返回值是 List
     * 2. List 不为空
     * 3. List 元素不是 List
     * 4. List 元素的类中有字段标注了 @FlattenList 注解
     *
     * @param obj 返回对象
     * @return boolean
     */
    @Override
    public boolean support(Object obj) {
        if (!(obj instanceof List)) {
            throw new ExcelException("@ResponseExcel 返回值必须为List类型");
        }

        List<?> objList = (List<?>) obj;

        // 空 List 或元素是 List，交给其他 Handler 处理
        if (objList.isEmpty() || objList.get(0) instanceof List) {
            return false;
        }

        // 检查是否有 @FlattenList 注解
        Class<?> dataClass = objList.get(0).getClass();
        return hasFlattenListAnnotation(dataClass);
    }

    @Override
    @SuppressWarnings("unchecked")
    public void write(Object obj, HttpServletResponse response, ExportExcel responseExcel) {
        List<?> objList = (List<?>) obj;

        if (objList.isEmpty()) {
            throw new ExcelException("FlattenList 导出不支持空数据");
        }

        Class<?> dataClass = objList.get(0).getClass();

        // 1. 展开数据
        List<Map<String, Object>> expandedData = ListEntityExpander.expandData((List<Object>) objList);

        // 2. 生成元数据
        ListEntityExpander.ListExpandMetadata metadata = ListEntityExpander.analyzeClass(dataClass);

        // 3. 生成合并区域
        List<ListEntityExpander.MergeRegion> mergeRegions =
            ListEntityExpander.generateMergeRegions(expandedData, metadata);

        // 4. 生成表头
        List<String> headers = ListEntityExpander.generateHeaders(metadata);
        List<List<String>> head = headers.stream()
            .map(Collections::singletonList)
            .collect(Collectors.toList());

        // 5. 转换数据为 List<List<Object>>
        List<List<Object>> rows = ListEntityExpander.convertToListData(expandedData, headers);

        // 获取 ExcelWriter
        ExcelWriter excelWriter = getExcelWriter(response, responseExcel);

        try {
            // 创建 WriteSheet（使用自定义表头）
            WriteSheet writeSheet = EasyExcel.writerSheet(
                responseExcel.sheets()[0].sheetName()
            )
            .head(head)
            .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
            .registerWriteHandler(new MergeCellHandler(mergeRegions))
            .build();

            // 写入数据
            excelWriter.write(rows, writeSheet);
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * 检查类中是否有字段标注了 @FlattenList 注解
     *
     * @param clazz 数据类
     * @return true 有注解，false 无注解
     */
    private boolean hasFlattenListAnnotation(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            FlattenList flattenList = field.getAnnotation(FlattenList.class);
            if (flattenList != null) {
                return true;
            }
        }
        return false;
    }

    /**
     * 合并单元格处理器
     */
    private static class MergeCellHandler implements com.alibaba.excel.write.handler.SheetWriteHandler {
        private final List<ListEntityExpander.MergeRegion> mergeRegions;

        public MergeCellHandler(List<ListEntityExpander.MergeRegion> mergeRegions) {
            this.mergeRegions = mergeRegions;
        }

        public void afterSheetCreate(com.alibaba.excel.write.metadata.holder.WriteSheetHolder writeSheetHolder,
                                     com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder writeWorkbookHolder) {
            Sheet sheet = writeSheetHolder.getSheet();

            // 应用合并区域
            for (ListEntityExpander.MergeRegion region : mergeRegions) {
                // 注意：EasyExcel 的行索引从0开始，但第0行是表头，所以数据从第1行开始
                // 需要将 region 的行号 +1
                CellRangeAddress cellRangeAddress = new CellRangeAddress(
                    region.getFirstRow() + 1,  // +1 跳过表头
                    region.getLastRow() + 1,
                    region.getFirstColumn(),
                    region.getLastColumn()
                );
                sheet.addMergedRegion(cellRangeAddress);
            }
        }
    }
}
