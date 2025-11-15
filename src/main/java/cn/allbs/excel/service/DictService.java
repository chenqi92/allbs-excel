package cn.allbs.excel.service;

/**
 * 字典服务接口
 * <p>
 * 用于提供字典数据的查询服务，用户需要实现此接口并注册为 Spring Bean
 * </p>
 *
 * <p>实现示例：</p>
 * <pre>
 * &#64;Service
 * public class DictServiceImpl implements DictService {
 *     &#64;Override
 *     public String getLabel(String dictType, String dictValue) {
 *         // 从数据库或缓存中查询字典标签
 *         return dictMapper.selectLabelByTypeAndValue(dictType, dictValue);
 *     }
 *
 *     &#64;Override
 *     public String getValue(String dictType, String dictLabel) {
 *         // 从数据库或缓存中查询字典值
 *         return dictMapper.selectValueByTypeAndLabel(dictType, dictLabel);
 *     }
 * }
 * </pre>
 *
 * @author ChenQi
 * @since 2025-11-15
 */
public interface DictService {

    /**
     * 根据字典类型和字典值获取字典标签
     * <p>
     * 用于导出时将字典值转换为字典标签
     * </p>
     *
     * @param dictType  字典类型，如：sys_user_sex
     * @param dictValue 字典值，如：1
     * @return 字典标签，如：男
     */
    String getLabel(String dictType, String dictValue);

    /**
     * 根据字典类型和字典标签获取字典值
     * <p>
     * 用于导入时将字典标签转换为字典值
     * </p>
     *
     * @param dictType  字典类型，如：sys_user_sex
     * @param dictLabel 字典标签，如：男
     * @return 字典值，如：1
     */
    String getValue(String dictType, String dictLabel);
}

