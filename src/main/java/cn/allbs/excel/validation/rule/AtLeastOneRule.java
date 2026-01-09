package cn.allbs.excel.validation.rule;

import cn.allbs.excel.annotation.cross.AtLeastOne;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.vo.FieldError;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 至少一个非空校验规则实现
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class AtLeastOneRule implements CrossValidationRule {

    private final AtLeastOne annotation;

    public AtLeastOneRule(AtLeastOne annotation) {
        this.annotation = annotation;
    }

    @Override
    public List<FieldError> validate(Object target, Class<?>... groups) {
        List<FieldError> errors = new ArrayList<>();

        if (!matchGroups(groups)) {
            return errors;
        }

        String[] fields = annotation.fields();
        int min = annotation.min();
        int filledCount = 0;

        for (String fieldName : fields) {
            Object value = FieldAccessorCache.getFieldValue(target, fieldName);
            if (ValidationHelper.isNotEmpty(value)) {
                filledCount++;
            }
        }

        if (filledCount < min) {
            String fieldNames = Arrays.stream(fields)
                    .map(f -> ValidationHelper.getExcelFieldName(target.getClass(), f))
                    .collect(Collectors.joining("、"));

            errors.add(FieldError.builder()
                    .fieldName(fieldNames)
                    .propertyName(String.join(",", fields))
                    .errorType("AtLeastOne")
                    .message(annotation.message())
                    .fullMessage("【" + fieldNames + "】" + annotation.message())
                    .build());
        }

        return errors;
    }

    @Override
    public Class<?>[] getGroups() {
        return annotation.groups();
    }
}

