package org.cts.hybrid.tests;

import java.util.Map;

import org.cts.hybrid.ExcelReader.ExcelDataProvider;
import org.cts.hybrid.annotation.ExcelDetails;
import org.testng.annotations.Test;

@ExcelDetails
public class ReadExcelData extends ExcelDataProvider {

	@Test(dataProvider = "data")
	public void readExcelData_Test(Map<String, String> input) {
		System.out.println(input.get("Name"));
	}



}
