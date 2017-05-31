package cn.logow.util.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.METHOD})
public @interface Column {

	String title() default "";
	
	int width() default 0;
	
	Align align() default Align.DEFAULT;
	
	String format() default "";
	
	int order() default Integer.MAX_VALUE;
	
	public enum Align {
		LEFT, CENTER, RIGHT, DEFAULT
	}
}
