package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.FieldsMatch;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.vo.FieldError;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * 字段值匹配校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class FieldsMatchRule implements CrossValidationRule {

    private final String fieldName;
    private final FieldsMatch annotation;

    public FieldsMatchRule(String fieldName, FieldsMatch annotation) {
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
        Object matchValue = FieldAccessorCache.getFieldValue(target, annotation.field());

        // 如果两个都为空，视为匹配
        if (ValidationHelper.isEmpty(currentValue) && ValidationHelper.isEmpty(matchValue)) {
            return errors;
        }

        // 检查是否匹配
        boolean matched;
        if (annotation.ignoreCase() && currentValue instanceof String && matchValue instanceof String) {
            matched = ((String) currentValue).equalsIgnoreCase((String) matchValue);
        } else {
            matched = Objects.equals(currentValue, matchValue);
        }

        if (!matched) {
            String excelFieldName = ValidationHelper.getExcelFieldName(target.getClass(), fieldName);
            String matchExcelFieldName = ValidationHelper.getExcelFieldName(target.getClass(), annotation.field());

            errors.add(FieldError.builder()
                    .fieldName(excelFieldName)
                    .propertyName(fieldName)
                    .errorType("FieldsMatch")
                    .message(annotation.message())
                    .fullMessage("【" + excelFieldName + "】与【" + matchExcelFieldName + "】" + annotation.message())
                    .build());
        }

        return errors;
    }

    @Override
    public Class<?>[] getGroups() {
        return annotation.groups();
    }
}

