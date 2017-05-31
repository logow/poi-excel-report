package cn.logow.util.excel;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.Date;

import cn.logow.util.excel.Column.Align;

class ColumnSpec {
	
	static final String FORMAT_DATE = "yyyy-mm-dd";
	static final String FORMAT_DECIMAL = "#,##0.00";
	
	String name;
	String title;
	int width;
	Align align;
	String format;
	Field field;
	Method getter;
	int order;
	
	ColumnSpec(String name, String title) {
		this.name = name;
		this.title = title != null && title.length() > 0 ? title : name;
	}
	
	static ColumnSpec createSpec(String name, Column col) {
		ColumnSpec spec = new ColumnSpec(name, col.title());
		spec.width = col.width();
		spec.align = col.align();
		spec.format = col.format();
		spec.order = col.order();
		return spec;
	}
	
	static ColumnSpec createSpec(String name, String title, String width, String align, String format) {
		ColumnSpec spec = new ColumnSpec(name, title);
		spec.width = 0;
		if (width != null) {
			try {
				spec.width = Integer.parseInt(width);
			} catch (NumberFormatException ignore) {
			}
		}
		if (spec.width > 255) {
			spec.width = 255;
		}
		
		spec.align = Align.DEFAULT;
		if (align != null) {
			try {
				spec.align = Align.valueOf(align.toUpperCase());
			} catch (Exception ignore) {
			}
		}

		spec.format = format;
		
		return spec;
	}
	
	void setDefaultFormat() {
		if (format != null && format.length() > 0) {
			return;
		}
		if (field == null && getter == null) {
			return;
		}
		
		Class<?> type = field != null ? field.getType() : getter.getReturnType();
		if (type == String.class) {
			return;
		}
		if (type.isPrimitive()) {
			if (type == float.class || type == double.class) {
				format = ColumnSpec.FORMAT_DECIMAL;
			}
			return;
		}
		if (Number.class.isAssignableFrom(type)) {
			if (type == Float.class || type == Double.class || type == BigDecimal.class) {
				format = ColumnSpec.FORMAT_DECIMAL;
			}
			return;
		}
		if (Date.class.isAssignableFrom(type)) {
			format = ColumnSpec.FORMAT_DATE;
			return;
		}
	}

	@Override
	public String toString() {
		return "Column [name=" + name + ", title=" + title + ", width=" + width + ", align=" + align + ", format="
				+ format + "]";
	}
}