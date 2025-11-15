package cn.allbs.excel.handle;

import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.enhance.WriterBuilderEnhancer;
import cn.allbs.excel.kit.ExcelException;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.util.CollectionUtils;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
public class SingleSheetWriteHandler extends AbstractSheetWriteHandler {
    public SingleSheetWriteHandler(ExcelConfigProperties configProperties,
                                   ObjectProvider<List<Converter<?>>> converterProvider, WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * obj 是List 且（list为空 或 list中的元素不是List）才返回true
     * 支持空List导出只有表头的Excel
     *
     * @param obj 返回对象
     * @return boolean
     */
    @Override
    public boolean support(Object obj) {
        if (obj instanceof List) {
            List<?> objList = (List<?>) obj;
            // 支持空List或者非嵌套List
            return objList.isEmpty() || !(objList.get(0) instanceof List);
        } else {
            throw new ExcelException("@ResponseExcel 返回值必须为List类型");
        }
    }

    @Override
    public void write(Object obj, HttpServletResponse response, ExportExcel responseExcel) {
        List<?> eleList = (List<?>) obj;
        ExcelWriter excelWriter = getExcelWriter(response, responseExcel);

        WriteSheet sheet;
        int totalRows = eleList != null ? eleList.size() : 0;
        if (CollectionUtils.isEmpty(eleList)) {
            // 空数据时，尝试从注解中获取数据类型
            Class<?> clazz = responseExcel.sheets()[0].clazz();
            if (clazz != Void.class) {
                // 如果指定了数据类型，使用该类型生成表头
                sheet = this.sheet(responseExcel.sheets()[0], clazz, responseExcel.template(),
                        responseExcel.headGenerator(), responseExcel.onlyExcelProperty(), responseExcel.autoMerge(), totalRows);
            } else {
                // 未指定数据类型，只创建空sheet（无表头）
                sheet = EasyExcel.writerSheet(responseExcel.sheets()[0].sheetName()).build();
            }
        } else {
            // 有数据时，从第一个元素获取类型
            Class<?> dataClass = eleList.get(0).getClass();
            sheet = this.sheet(responseExcel.sheets()[0], dataClass, responseExcel.template(),
                    responseExcel.headGenerator(), responseExcel.onlyExcelProperty(), responseExcel.autoMerge(), totalRows);
        }

        // 填充 sheet
        if (responseExcel.fill()) {
            excelWriter.fill(eleList, sheet);
        } else {
            // 写入sheet
            excelWriter.write(eleList, sheet);
        }
        excelWriter.finish();
    }
}
