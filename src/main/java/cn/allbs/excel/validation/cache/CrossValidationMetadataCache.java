package cn.allbs.excel.validation.cache;

import cn.allbs.excel.annotation.cross.*;
import cn.allbs.excel.validation.rule.*;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 跨字段校验元数据缓存
 * <p>
 * 缓存类的注解解析结果，避免每行数据都重复解析。
 * </p>
 * <p>
 * <b>注意：</b>缓存使用 static ConcurrentHashMap 实现，适用于 DTO 类数量有限的场景。
 * 在热部署或动态类加载场景下，请在适当时机调用 {@link #clearCache()} 避免内存泄漏。
 * </p>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public class CrossValidationMetadataCache {

    private static final Map<Class<?>, CrossValidationMetadata> CACHE = new ConcurrentHashMap<>();

    private CrossValidationMetadataCache() {
    }

    /**
     * 获取或解析类的校验元数据
     *
     * @param clazz 类
     * @return 校验元数据
     */
    public static CrossValidationMetadata getOrParse(Class<?> clazz) {
        return CACHE.computeIfAbsent(clazz, CrossValidationMetadataCache::parse);
    }

    /**
     * 解析类的校验元数据
     */
    private static CrossValidationMetadata parse(Class<?> clazz) {
        CrossValidationMetadata metadata = new CrossValidationMetadata();

        // 解析字段级别注解（包括父类字段）
        Class<?> currentClass = clazz;
        while (currentClass != null && currentClass != Object.class) {
            for (Field field : currentClass.getDeclaredFields()) {
                parseFieldAnnotations(field, metadata);
            }
            currentClass = currentClass.getSuperclass();
        }

        // 解析类级别注解（包括父类）
        currentClass = clazz;
        while (currentClass != null && currentClass != Object.class) {
            parseClassAnnotations(currentClass, metadata);
            currentClass = currentClass.getSuperclass();
        }

        return metadata;
    }

    /**
     * 解析字段级别注解
     */
    private static void parseFieldAnnotations(Field field, CrossValidationMetadata metadata) {
        String fieldName = field.getName();

        // @RequiredIf
        RequiredIf[] requiredIfs = field.getAnnotationsByType(RequiredIf.class);
        for (RequiredIf annotation : requiredIfs) {
            metadata.addRule(new RequiredIfRule(fieldName, annotation));
        }

        // @MutualExclusive
        MutualExclusive mutualExclusive = field.getAnnotation(MutualExclusive.class);
        if (mutualExclusive != null) {
            metadata.addRule(new MutualExclusiveRule(fieldName, mutualExclusive));
        }

        // @FieldsMatch
        FieldsMatch fieldsMatch = field.getAnnotation(FieldsMatch.class);
        if (fieldsMatch != null) {
            metadata.addRule(new FieldsMatchRule(fieldName, fieldsMatch));
        }

        // @AllowedValuesIf
        AllowedValuesIf[] allowedValuesIfs = field.getAnnotationsByType(AllowedValuesIf.class);
        for (AllowedValuesIf annotation : allowedValuesIfs) {
            metadata.addRule(new AllowedValuesIfRule(fieldName, annotation));
        }

        // @ConditionalPattern
        ConditionalPattern conditionalPattern = field.getAnnotation(ConditionalPattern.class);
        if (conditionalPattern != null) {
            metadata.addRule(new ConditionalPatternRule(fieldName, conditionalPattern));
        }
    }

    /**
     * 解析类级别注解
     */
    private static void parseClassAnnotations(Class<?> clazz, CrossValidationMetadata metadata) {
        // @AtLeastOne
        AtLeastOne[] atLeastOnes = clazz.getAnnotationsByType(AtLeastOne.class);
        for (AtLeastOne annotation : atLeastOnes) {
            metadata.addRule(new AtLeastOneRule(annotation));
        }

        // @FieldsCompare
        FieldsCompare[] fieldsCompares = clazz.getAnnotationsByType(FieldsCompare.class);
        for (FieldsCompare annotation : fieldsCompares) {
            metadata.addRule(new FieldsCompareRule(annotation));
        }

        // @FieldsCalculation
        FieldsCalculation[] fieldsCalculations = clazz.getAnnotationsByType(FieldsCalculation.class);
        for (FieldsCalculation annotation : fieldsCalculations) {
            metadata.addRule(new FieldsCalculationRule(annotation));
        }

        // @UniqueCombination - 需要特殊处理，在 CrossFieldValidators 中处理
        UniqueCombination[] uniqueCombinations = clazz.getAnnotationsByType(UniqueCombination.class);
        for (UniqueCombination annotation : uniqueCombinations) {
            metadata.addUniqueCombination(annotation);
        }

        // @CrossFieldExpression
        CrossFieldExpression[] expressions = clazz.getAnnotationsByType(CrossFieldExpression.class);
        for (CrossFieldExpression annotation : expressions) {
            metadata.addRule(new CrossFieldExpressionRule(annotation));
        }
    }

    /**
     * 清除缓存（用于测试）
     */
    public static void clearCache() {
        CACHE.clear();
    }

    /**
     * 校验元数据
     */
    public static class CrossValidationMetadata {
        private final List<CrossValidationRule> rules = new ArrayList<>();
        private final List<UniqueCombination> uniqueCombinations = new ArrayList<>();

        public void addRule(CrossValidationRule rule) {
            rules.add(rule);
        }

        public void addUniqueCombination(UniqueCombination annotation) {
            uniqueCombinations.add(annotation);
        }

        public List<CrossValidationRule> getRules() {
            return rules;
        }

        public List<UniqueCombination> getUniqueCombinations() {
            return uniqueCombinations;
        }

        public boolean hasRules() {
            return !rules.isEmpty() || !uniqueCombinations.isEmpty();
        }
    }
}
