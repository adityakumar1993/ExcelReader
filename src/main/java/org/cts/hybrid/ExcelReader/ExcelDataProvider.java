package org.cts.hybrid.ExcelReader;

import java.lang.annotation.Annotation;

import org.cts.hybrid.annotation.ExcelDetails;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;

public class ExcelDataProvider {

	private String excelName;
	private String sheetName;

	@BeforeClass
	public void getExcelDetails() {
		Class<?> obj = getClass();
		if (obj.isAnnotationPresent(ExcelDetails.class)) {
			Annotation annotation = obj.getAnnotation(ExcelDetails.class);
			ExcelDetails excelDetails = (ExcelDetails) annotation;
			if (excelDetails.excelName().equals("")) {
				excelName = getClass().getSimpleName();
			} else {
				excelName = excelDetails.excelName();
			}
			sheetName = excelDetails.sheetName();
		}
	}

	@DataProvider
	public Object[][] data() {
		return ReadExcel.readData(excelName, sheetName);
	}
}
