package cn.logow.util.excel.test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import cn.logow.util.excel.ExcelReport;

public class ExcelReportTest {

	public static void main(String[] args) throws IOException {
		ExcelReport<User> ee = ExcelReport.xlsx(User.class);
		ee.setRowsPerSheet(30);
		List<User> users = new ArrayList<User>();
		for (int i = 1; i < 100; i++) {
			users.add(new User("user" + i, 20 + i, 10000 + i * 100));
		}
		
		OutputStream out = new FileOutputStream("D:/tmp/export7.xlsx");
		ee.addAll(users);
		ee.exportTo(out);
		out.close();
	}
}
