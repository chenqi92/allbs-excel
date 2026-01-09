package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.MutualExclusive;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.vo.FieldError;

import java.util.ArrayList;
import java.util.List;

/**
 * 互斥校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class MutualExclusiveRule implements CrossValidationRule {

    private final String fieldName;
    private final MutualExclusive annotation;

    public MutualExclusiveRule(String fieldName, MutualExclusive annotation) {
        this.fieldName = fieldName;
        this.annotation = annotation;
    }

    @Override
    public List<FieldError> validate(Object target, Class<?>... groups) {
        List<FieldError> errors = new ArrayList<>();

        if (!matchGroups(groups)) {
            return errors;
        }

        Object currentValue = FieldAccessorCache.getFieldValue(target, fieldName);

        // 如果当前字段为空，跳过校验
        if (ValidationHelper.isEmpty(currentValue)) {
            return errors;
        }

        // 检查互斥字段是否有值
        String[] mutualFields = annotation.fields();
        for (String mutualField : mutualFields) {
            Object mutualValue = FieldAccessorCache.getFieldValue(target, mutualField);
            if (ValidationHelper.isNotEmpty(mutualValue)) {
                String excelFieldName = ValidationHelper.getExcelFieldName(target.getClass(), fieldName);
                String conflictExcelFieldName = ValidationHelper.getExcelFieldName(target.getClass(), mutualField);

                errors.add(FieldError.builder()
                        .fieldName(excelFieldName)
                        .propertyName(fieldName)
                        .errorType("MutualExclusive")
                        .message(annotation.message())
                        .fullMessage("【" + excelFieldName + "】与【" + conflictExcelFieldName + "】" + annotation.message())
                        .build());
                break; // 只报告第一个冲突
            }
        }

        return errors;
    }

    @Override
    public Class<?>[] getGroups() {
        return annotation.groups();
    }
}

