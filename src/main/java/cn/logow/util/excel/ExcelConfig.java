package cn.logow.util.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import cn.logow.util.excel.Column.Align;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelConfig {
	
	Align titleAlign() default Align.CENTER;
	
	Align columnAlign() default Align.CENTER;
	
	String sheetPrefix() default "";
	
	int rowsPerSheet() default 0;
	
	int rowHeight() default 0;
	
	int columnWidth() default 0;
}
