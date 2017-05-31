package cn.logow.util.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReport<E> {
	
	static final String XLS_FORMAT = "xls";
	static final String XLSX_FORMAT = "xlsx";
	
	static final String DEFAULT_DATA_FORMAT = XLS_FORMAT;
	
	static final int DEFAULT_BUF_SIZE = 100;
	
	static final int DEFAULT_ROWS_PER_SHEET = 10000;
	
	static final Log log = LogFactory.getLog(ExcelReport.class);
	
	private Workbook workbook;
	private Sheet currentSheet;
	private RowRenderer<E> renderer;
	private List<E> buffer = new ArrayList<E>();
	private String sheetPrefix = "";
	
	private int bufSize = DEFAULT_BUF_SIZE;
	private int rowsPerSheet = DEFAULT_ROWS_PER_SHEET;
	private int defaultRowHeight = 0;
	
	public static <E> ExcelReport<E> xls(Class<E> entityClass) {
		return new ExcelReport<E>(entityClass, XLS_FORMAT);
	}
	
	public static <E> ExcelReport<E> xlsx(Class<E> entityClass) {
		return new ExcelReport<E>(entityClass, XLSX_FORMAT);
	}
	
	public static <E> ExcelReport<E> create(Class<E> entityClass) {
		return new ExcelReport<E>(entityClass, DEFAULT_DATA_FORMAT);
	}
	
	protected ExcelReport(Class<E> entityClass, String dataFormat) {
		workbook = createWorkbook(dataFormat);
		renderer = createRenderer(entityClass, workbook);
		currentSheet = createNewSheet();
		ExcelConfig config = entityClass.getAnnotation(ExcelConfig.class);
		if (config != null) {
			defaultRowHeight = config.rowHeight();
			sheetPrefix = config.sheetPrefix();
			if (config.rowsPerSheet() > 0) {
				setRowsPerSheet(config.rowsPerSheet());
			}
		}
	}
	
	public void setRowsPerSheet(int maxRows) {
		if (maxRows <= 0) {
			throw new IllegalArgumentException("rowsPerSheet must be positive");
		}
		rowsPerSheet = maxRows;
	}
	
	public void exportTo(OutputStream out) throws IOException {
		long t1 = System.currentTimeMillis();
		int sheets = workbook.getNumberOfSheets();
		ColumnSpec[] columns = renderer.getColumns();
		for (int i = 0; i < sheets; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			for (int j = 0; j < columns.length; j++) {
				if (columns[j].width <= 0) {
					sheet.autoSizeColumn(j);
				}
			}
		}
		long t2 = System.currentTimeMillis();
		workbook.write(out);
		out.flush();
		long t3 = System.currentTimeMillis();
		
		if (log.isDebugEnabled()) {
			log.debug("auto size columns cost: " + (t2 - t1) + "ms");
			log.debug("export excel cost: " + (t3 - t2) + "ms");
		}
	}
	
	public void addAll(List<E> list) {
		if (!buffer.isEmpty()) {
			flush();
		}
		addRows(list);
	}
	
	public void add(E e) {
		buffer.add(e);
		if (buffer.size() >= bufSize) {
			flush();
		}
	}
	
	public void flush() {
		addRows(buffer);
		buffer.clear();
	}
	
	private void addRows(List<E> list) {
		int leftRows = list.size();
		int addedRows = 0;
		while (leftRows > 0) {
			int sheetRows = currentSheet.getLastRowNum() + 1;
			if (sheetRows >= rowsPerSheet) {
				currentSheet = createNewSheet();
				sheetRows = currentSheet.getLastRowNum() + 1;
			}
			int availRows = rowsPerSheet - sheetRows;
			int rowsToAdd = Math.min(availRows, leftRows);
			addRowsToSheet(list.subList(addedRows, addedRows + rowsToAdd));
			addedRows += rowsToAdd;
			leftRows -= rowsToAdd;
		}
	}
	
	private Workbook createWorkbook(String format) {
		if (XLS_FORMAT.equalsIgnoreCase(format)) {
			return new HSSFWorkbook();
		} else if (XLSX_FORMAT.equalsIgnoreCase(format)) {
			return new XSSFWorkbook();
		} else {
			throw new IllegalArgumentException("invalid excel format: " + format);
		}
	}
	
	private RowRenderer<E> createRenderer(Class<E> entityClass, Workbook workbook) {
		return new RowRenderer<E>(entityClass, workbook);
	}
	
	private Sheet createNewSheet() {
		String sheetName = getSheetName();
		Sheet sheet = sheetName != null ? workbook.createSheet(sheetName) : workbook.createSheet();
		ColumnSpec[] columns = renderer.getColumns();
		for (int i = 0; i < columns.length; i++) {
			if (columns[i].width > 0) {
				sheet.setColumnWidth(i, columns[i].width * 256);
			}
		}
		
		if (defaultRowHeight > 0) {
			sheet.setDefaultRowHeight((short) defaultRowHeight);
		}
		
		addTitleRow(sheet, columns);
		
		return sheet;
	}
	
	private String getSheetName() {
		if (sheetPrefix.isEmpty()) {
			return null;
		}
		
		String sheetName = sheetPrefix + workbook.getNumberOfSheets();
		int len = sheetName.length();
		if (len >= 31) {
			return sheetName.substring(len - 31);
		} else {
			return sheetName;
		}
	}

	private void addRowsToSheet(List<E> rows) {
		try {
			int rowNum = currentSheet.getLastRowNum();
			for (E data : rows) {
				Row row = currentSheet.createRow(++rowNum);
				renderer.renderCell(row, data);
			}
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}
	
	private void addTitleRow(Sheet sheet, ColumnSpec[] columns) {
		Row title = sheet.createRow(0);
		title.setHeight((short) (sheet.getDefaultRowHeight() * 1.5));
		renderer.renderTitle(title);
		sheet.createFreezePane( 0, 1, 0, 1 ); //冻结标题行
	}
}
