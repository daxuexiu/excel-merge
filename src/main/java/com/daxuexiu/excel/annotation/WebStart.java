package com.daxuexiu.excel.annotation;

import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.lang.annotation.*;

/**
 * @Author yueylong
 * @Date 19-7-9 下午2:10
 * @Desc
 **/
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
@SpringBootApplication
public @interface WebStart {
}
