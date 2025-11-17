package cn.allbs.excel.head;

import cn.allbs.excel.util.FlattenFieldProcessor;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * 支持 @FlattenProperty 的表头生成器
 * <p>
 * 自动展开嵌套对象中标注了 @ExcelProperty 的字段
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-17
 */
@Slf4j
public class FlattenHeadGenerator implements HeadGenerator {

    @Override
    public HeadMeta head(Class<?> clazz) {
        HeadMeta headMeta = new HeadMeta();

        // 处理展开字段
        List<FlattenFieldProcessor.FlattenFieldInfo> fieldInfos =
                FlattenFieldProcessor.processFlattenFields(clazz);

        // 生成表头列表
        List<List<String>> headList = new ArrayList<>();
        Set<String> ignoreFields = new HashSet<>();

        for (FlattenFieldProcessor.FlattenFieldInfo info : fieldInfos) {
            // 添加表头
            headList.add(Collections.singletonList(info.getHeadName()));

            // 如果是展开字段，需要忽略原始的嵌套对象字段
            if (info.isFlatten() && info.getParentField() != null) {
                String parentFieldName = info.getParentField().getName();
                ignoreFields.add(parentFieldName);
            }
        }

        headMeta.setHead(headList);
        headMeta.setIgnoreHeadFields(ignoreFields);

        log.debug("Generated flatten head: {} ignore fields: {}", headList, ignoreFields);

        return headMeta;
    }
}
