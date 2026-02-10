package cn.allbs.excel.validation.rule;

import cn.allbs.excel.vo.FieldError;

import java.util.List;

/**
 * 跨字段校验规则接口
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public interface CrossValidationRule {

    /**
     * 执行校验
     *
     * @param target 被校验的对象
     * @param groups 当前校验分组
     * @return 错误列表，为空表示校验通过
     */
    List<FieldError> validate(Object target, Class<?>... groups);

    /**
     * 获取校验分组
     *
     * @return 分组类数组
     */
    Class<?>[] getGroups();

    /**
     * 检查是否匹配指定分组
     *
     * @param groups 当前校验分组
     * @return 是否匹配
     */
    default boolean matchGroups(Class<?>... groups) {
        Class<?>[] ruleGroups = getGroups();

        // 如果规则没有指定分组，则始终匹配
        if (ruleGroups == null || ruleGroups.length == 0) {
            return true;
        }

        // 如果没有指定当前分组，则只匹配没有分组的规则（已在上面处理）
        if (groups == null || groups.length == 0) {
            return false;
        }

        // 检查是否有交集
        for (Class<?> ruleGroup : ruleGroups) {
            for (Class<?> group : groups) {
                if (ruleGroup.isAssignableFrom(group)) {
                    return true;
                }
            }
        }

        return false;
    }
}
