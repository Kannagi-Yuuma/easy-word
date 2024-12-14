package com.yuuma.constants;

/**
 * 占位符正则表达式常量类
 */
public class PlaceholderRegexConstant {

    /** 普通占位符正则表达式 */
    public static final String NORMAL = "\\$\\{([^}]+)}";

    /** 双大括号占位符正则表达式 */
    public static final String DOUBLE_BRACE = "\\{\\{([^}]+)}}";
}
