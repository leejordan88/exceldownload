package org.excel.demo.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Excel {

    /**
     * 해더 명
     * default = "field"
     * @return
     */
    String name() default "field";

    /**
     * 좌우 넓이
     * default = 20 * 100 FIXEL
     * @return
     */
    int width() default 20;


}
