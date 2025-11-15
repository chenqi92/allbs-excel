package cn.allbs.excel.handle;

import cn.allbs.excel.config.ExcelConfigProperties;
import cn.allbs.excel.annotation.ExportExcel;
import cn.allbs.excel.annotation.Sheet;
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
public class ManySheetWriteHandler extends AbstractSheetWriteHandler {
    public ManySheetWriteHandler(ExcelConfigProperties configProperties,
                                 ObjectProvider<List<Converter<?>>> converterProvider, WriterBuilderEnhancer excelWriterBuilderEnhance) {
        super(configProperties, converterProvider, excelWriterBuilderEnhance);
    }

    /**
     * 当List不为空且List中的元素也是List 才返回true
     * 注意：空List会被SingleSheetWriteHandler处理
     *
     * @param obj 返回对象
     * @return boolean
     */
    @Override
    public boolean support(Object obj) {
        if (obj instanceof List) {
            List<?> objList = (List<?>) obj;
            // 只处理非空的嵌套List，空List交给SingleSheetWriteHandler处理
            return !objList.isEmpty() && objList.get(0) instanceof List;
        } else {
            throw new ExcelException("@ResponseExcel 返回值必须为List类型");
        }
    }

    @Override
    public void write(Object obj, HttpServletResponse response, ExportExcel responseExcel) {
        List<?> objList = (List<?>) obj;
        ExcelWriter excelWriter = getExcelWriter(response, responseExcel);

        Sheet[] sheets = responseExcel.sheets();
        WriteSheet sheet;
        for (int i = 0; i < sheets.length; i++) {
            List<?> eleList = (List<?>) objList.get(i);

            if (CollectionUtils.isEmpty(eleList)) {
                // 空数据时，尝试从注解中获取数据类型
                Class<?> clazz = responseExcel.sheets()[i].clazz();
                if (clazz != Void.class) {
                    // 如果指定了数据类型，使用该类型生成表头
                    sheet = this.sheet(responseExcel.sheets()[i], clazz, responseExcel.template(),
                            responseExcel.headGenerator(), responseExcel.onlyExcelProperty(), responseExcel.autoMerge());
                } else {
                    // 未指定数据类型，只创建空sheet（无表头）
                    sheet = EasyExcel.writerSheet(responseExcel.sheets()[i].sheetName()).build();
                }
            } else {
                // 有数据时，从第一个元素获取类型
                Class<?> dataClass = eleList.get(0).getClass();
                sheet = this.sheet(responseExcel.sheets()[i], dataClass, responseExcel.template(),
                        responseExcel.headGenerator(), responseExcel.onlyExcelProperty(), responseExcel.autoMerge());
            }

            // 填充 sheet
            if (responseExcel.fill()) {
                excelWriter.fill(eleList, sheet);
            } else {
                // 写入sheet
                excelWriter.write(eleList, sheet);
            }
        }
        excelWriter.finish();
    }
}
