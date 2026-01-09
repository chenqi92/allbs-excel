package cn.allbs.excel.validation;

import cn.allbs.excel.annotation.cross.UniqueCombination;
import cn.allbs.excel.util.ValidationHelper;
import cn.allbs.excel.validation.cache.CrossValidationMetadataCache;
import cn.allbs.excel.validation.cache.CrossValidationMetadataCache.CrossValidationMetadata;
import cn.allbs.excel.validation.cache.FieldAccessorCache;
import cn.allbs.excel.validation.rule.CrossValidationRule;
import cn.allbs.excel.vo.FieldError;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 跨字段校验执行器
 * <p>
 * 负责解析并执行对象上的跨字段校验注解
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class CrossFieldValidators {

    private CrossFieldValidators() {
    }

    /**
     * 校验对象（无分组）
     *
     * @param target 被校验的对象
     * @return 错误列表
     */
    public static List<FieldError> validate(Object target) {
        return validate(target, (Class<?>[]) null);
    }

    /**
     * 校验对象（指定分组）
     *
     * @param target 被校验的对象
     * @param groups 校验分组
     * @return 错误列表
     */
    public static List<FieldError> validate(Object target, Class<?>... groups) {
        if (target == null) {
            return Collections.emptyList();
        }

        List<FieldError> errors = new ArrayList<>();
        CrossValidationMetadata metadata = CrossValidationMetadataCache.getOrParse(target.getClass());

        if (!metadata.hasRules()) {
            return errors;
        }

        // 执行行内校验规则
        for (CrossValidationRule rule : metadata.getRules()) {
            if (groups == null || groups.length == 0 || rule.matchGroups(groups)) {
                errors.addAll(rule.validate(target, groups));
            }
        }

        return errors;
    }

    /**
     * 创建唯一性校验器实例
     * <p>
     * 用于跨行唯一性校验，需要在整个导入过程中维护状态
     * </p>
     *
     * @param clazz 校验的类
     * @return 唯一性校验器
     */
    public static UniqueValidator createUniqueValidator(Class<?> clazz) {
        CrossValidationMetadata metadata = CrossValidationMetadataCache.getOrParse(clazz);
        return new UniqueValidator(clazz, metadata.getUniqueCombinations());
    }

    /**
     * 跨行唯一性校验器
     */
    public static class UniqueValidator {
        private final Class<?> clazz;
        private final List<UniqueCombination> annotations;
        private final Map<String, Set<String>> seenCombinations;

        public UniqueValidator(Class<?> clazz, List<UniqueCombination> annotations) {
            this.clazz = clazz;
            this.annotations = annotations;
            this.seenCombinations = new HashMap<>();
            for (int i = 0; i < annotations.size(); i++) {
                seenCombinations.put("rule_" + i, new HashSet<>());
            }
        }

        /**
         * 检查行数据是否重复
         *
         * @param target   行数据对象
         * @param rowIndex 行号
         * @param groups   校验分组
         * @return 错误列表
         */
        public List<FieldError> validate(Object target, int rowIndex, Class<?>... groups) {
            List<FieldError> errors = new ArrayList<>();

            for (int i = 0; i < annotations.size(); i++) {
                UniqueCombination annotation = annotations.get(i);

                // 检查分组
                if (!ValidationHelper.matchGroups(annotation.groups(), groups)) {
                    continue;
                }

                String[] fields = annotation.fields();
                List<String> values = new ArrayList<>();
                boolean hasNull = false;

                for (String fieldName : fields) {
                    Object value = FieldAccessorCache.getFieldValue(target, fieldName);
                    if (FieldAccessorCache.isEmpty(value)) {
                        hasNull = true;
                        break;
                    }
                    String strValue = String.valueOf(value);
                    if (annotation.ignoreCase()) {
                        strValue = strValue.toLowerCase();
                    }
                    values.add(strValue);
                }

                // 如果忽略空值且存在空值，跳过校验
                if (hasNull && annotation.ignoreNull()) {
                    continue;
                }

                if (hasNull) {
                    continue; // 有空值无法构建组合键
                }

                String combinationKey = String.join("|", values);
                Set<String> seen = seenCombinations.get("rule_" + i);

                if (!seen.add(combinationKey)) {
                    // 重复
                    String fieldNames = Arrays.stream(fields)
                            .map(this::getExcelFieldName)
                            .collect(Collectors.joining("、"));

                    errors.add(FieldError.builder()
                            .fieldName(fieldNames)
                            .propertyName(String.join(",", fields))
                            .errorType("UniqueCombination")
                            .message(annotation.message())
                            .fullMessage("【" + fieldNames + "】" + annotation.message() + "（行号：" + rowIndex + "）")
                            .build());
                }
            }

            return errors;
        }

        private String getExcelFieldName(String fieldName) {
            return ValidationHelper.getExcelFieldName(clazz, fieldName);
        }

        /**
         * 重置校验器状态
         */
        public void reset() {
            seenCombinations.values().forEach(Set::clear);
        }

        /**
         * 检查是否有唯一性校验规则
         */
        public boolean hasRules() {
            return !annotations.isEmpty();
        }
    }
}
