package cn.logow.util.excel.test;

import java.math.BigDecimal;
import java.util.Date;

import cn.logow.util.excel.Column;
import cn.logow.util.excel.Column.Align;
import cn.logow.util.excel.ExcelConfig;

@ExcelConfig(columnWidth = 20)
public class User {
	@Column(title = "姓名")
	private String name;
	
	@Column(title = "年龄")
	private Integer age;
	
	@Column(title = "出生日期")
	private Date birthday;
	
	@Column(title = "薪水", align = Align.RIGHT)
	private BigDecimal salary;
	
	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	public BigDecimal getSalary() {
		return salary;
	}

	public void setSalary(BigDecimal salary) {
		this.salary = salary;
	}

	public User(String name, Integer age, double salary) {
		this.name = name;
		this.age = age;
		this.birthday = new Date();
		this.salary = new BigDecimal(salary);
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}
}
