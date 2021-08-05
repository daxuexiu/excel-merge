package com.daxuexiu.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Author yueylong
 * @Date 19-7-9 下午2:13
 * @Desc
 **/
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    String title() default "";

    int index();

    int rowIndex() default 0;
}
