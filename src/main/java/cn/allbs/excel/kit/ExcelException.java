package cn.allbs.excel.kit;


import cn.allbs.common.constant.AllbsCoreConstants;

/**
 * 功能:
 *
 * @author ChenQi
 * @version 1.0
 * @since 2021/3/29
 */
public class ExcelException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    public ExcelException(String message) {
        super(AllbsCoreConstants.ALLBS_TIP + message);
    }
}
