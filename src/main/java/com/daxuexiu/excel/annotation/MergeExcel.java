package com.daxuexiu.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Author yueylong
 * @Date 19-7-9 下午2:13
 * @Desc
 */
@Target({ElementType.FIELD,ElementType.TYPE,ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface MergeExcel {
    /**
     * 单元格内的值
     *
     * @return
     */
    String value() default "";
    /**
     * 合并行的数量
     *
     * @return
     */
    int rowspan() default 0;
    /**
     * 起始列
     *
     * @return
     */
    int colIndex();
    /**
     * 列合并的数量
     *
     * @return
     */
    int colspan() default 0;
}
