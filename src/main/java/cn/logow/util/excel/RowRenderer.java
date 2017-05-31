package cn.logow.util.excel;

import java.util.Date;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import cn.logow.util.excel.Column.Align;

class RowRenderer<E> {
	
	static final Log log = LogFactory.getLog(RowRenderer.class);
	
	private ColumnSpec[] specs;
	private CellStyle[] styles;
	
	RowRenderer(Class<?> entityClass, Workbook wb) {
		specs = createSpecs(entityClass);
		styles = createStyles(specs, wb);
	}
	
	ColumnSpec[] getColumns() {
		return specs;
	}
	
	void renderCell(Row row, E entity) throws Exception {
		for (int i = 0; i < specs.length; i++) {
			Object val = null;
			if (specs[i].field != null) {
				val = specs[i].field.get(entity);
			} else {
				val = specs[i].getter.invoke(entity);
			}
			if (val != null) {
				Cell cell = row.createCell(i);
				cell.setCellStyle(styles[i]);
				if (val instanceof String) {
					cell.setCellValue(val.toString());
				} else if (val instanceof Number) {
					cell.setCellValue(((Number) val).doubleValue());
				} else if (val instanceof Date){
					cell.setCellValue((Date) val);
				} else {
					cell.setCellValue(val.toString());
				}
			}
		}
	}
	
	void renderTitle(Row row) {
		CellStyle cs = styles[specs.length];
		for (int i = 0; i < specs.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(specs[i].title);
			cell.setCellStyle(cs);
		}
	}
	
	private ColumnSpec[] createSpecs(Class<?> entityClass) {
		ColumnSpec[] specs = MapperResolver.resolveColumns(entityClass);
		if (specs == null || specs.length == 0) {
			throw new IllegalArgumentException("No column mappings for entity class: " + entityClass);
		}
		
		ExcelConfig config = entityClass.getAnnotation(ExcelConfig.class);
		if (config != null) {
			for (ColumnSpec spec : specs) {
				if (spec.align == null || spec.align == Align.DEFAULT) {
					spec.align = config.columnAlign();
				}
				if (spec.width <= 0) {
					spec.width = config.columnWidth();
				}
			}
		}
		
		if (log.isDebugEnabled()) {
			log.debug("Actual specs for entity [" + entityClass.getName() + "]");
			for (ColumnSpec spec : specs) {
				log.debug(spec);
			}
		}
		
		return specs;
	}
	
	private CellStyle[] createStyles(ColumnSpec[] specs, Workbook wb) {
		CellStyle[] styles = new CellStyle[specs.length + 1];
		
		Font cellFont = wb.createFont();
		cellFont.setFontName("微软雅黑");
		for (int i = 0; i < specs.length; i++) {
			styles[i] = createCellStyle(specs[i], cellFont, wb);
		}
		
		styles[specs.length] = createTitleStyle(wb);
		
		return styles;
	}
	
	private CellStyle createTitleStyle(Workbook wb) {
		Font titleFont = wb.createFont();
		titleFont.setFontName("宋体");
		titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		CellStyle cs = wb.createCellStyle();
		cs.setAlignment(CellStyle.ALIGN_CENTER);
		cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cs.setWrapText(true);
		cs.setFont(titleFont);
		return cs;
	}
	
	private CellStyle createCellStyle(ColumnSpec spec, Font cellFont, Workbook wb) {
		CellStyle style = wb.createCellStyle();
		style.setFont(cellFont);

		switch(spec.align) {
			case LEFT: style.setAlignment(CellStyle.ALIGN_LEFT); break;
			case CENTER: style.setAlignment(CellStyle.ALIGN_CENTER); break;
			case RIGHT: style.setAlignment(CellStyle.ALIGN_RIGHT); break;
			default:
		}

		if (spec.format != null && spec.format.length() > 0) {
			short df = wb.createDataFormat().getFormat(spec.format);
			style.setDataFormat(df);
		}
		
		return style;
	}
}
