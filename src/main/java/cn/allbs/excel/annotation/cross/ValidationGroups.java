package cn.allbs.excel.annotation.cross;

/**
 * 预定义校验分组接口
 * <p>
 * 类似 JSR-380 的分组校验，允许在不同场景下应用不同的校验规则
 * </p>
 *
 * <p>使用示例：</p>
 * <pre>
 * &#64;RequiredIf(
 *     field = "employeeType",
 *     hasValue = "正式员工",
 *     groups = {ValidationGroups.Submit.class},
 *     message = "正式员工必须填写工号"
 * )
 * private String employeeId;
 * </pre>
 *
 * @author ChenQi
 * @since 2026-01-09
 */
public interface ValidationGroups {

    /**
     * 默认分组 - 始终执行的校验
     */
    interface Default {
    }

    /**
     * 创建时校验
     */
    interface Create {
    }

    /**
     * 更新时校验
     */
    interface Update {
    }

    /**
     * 草稿保存校验（宽松模式）
     */
    interface Draft {
    }

    /**
     * 提交审核校验（严格模式）
     */
    interface Submit {
    }

    /**
     * 导入校验
     */
    interface Import {
    }

    /**
     * 导出校验
     */
    interface Export {
    }
}
