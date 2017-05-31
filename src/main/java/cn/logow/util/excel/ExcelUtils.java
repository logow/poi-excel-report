package cn.logow.util.excel;

import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

public abstract class ExcelUtils {

	public static void download(ExcelReport<?> report, HttpServletResponse resp, String fileName) throws IOException {
		if (fileName == null) {
			throw new IllegalArgumentException("fileName must not be null");
		}
		resp.setContentType("application/octet-stream");
		resp.setHeader("Content-Disposition", "attachment; filename=" + new String(fileName.getBytes(), "ISO-8859-1"));
		report.flush();
		report.exportTo(resp.getOutputStream());
	}
}
