package org.cts.oneframework.excelreader;

import java.lang.annotation.Annotation;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

import org.cts.oneframework.annotation.ExcelDetails;

public class ExcelDataProvider {

	
	private String[] excelDetailsValue = new String[2];
	private Map<String, String[]> methodExcelDetails = new HashMap<>();
	private Class<?> obj = null;

	public ExcelDataProvider(Class<?> obj) {
		this.obj = obj;
	}

	private void getExcelDetailsFromClass() {
		if (obj.isAnnotationPresent(ExcelDetails.class)) {
			Annotation annotation = obj.getAnnotation(ExcelDetails.class);
			ExcelDetails excelDetails = (ExcelDetails) annotation;
			if (excelDetails.excelName().isEmpty()) {
				excelDetailsValue[0] = obj.getSimpleName();
			} else {
				excelDetailsValue[0] = excelDetails.excelName();
			}
			excelDetailsValue[1] = excelDetails.sheetName();
		}
	}

	private void getExcelDetailsFromMethod() {
		Method[] methodList = obj.getDeclaredMethods();
		for (Method method : methodList) {
			String[] excelInfo = new String[2];
			if (method.isAnnotationPresent(ExcelDetails.class)) {
				Annotation annotation = method.getAnnotation(ExcelDetails.class);
				ExcelDetails excelDetails = (ExcelDetails) annotation;
				if (excelDetails.excelName().isEmpty()) {
					excelInfo[0] = method.getName();
				} else {
					excelInfo[0] = excelDetails.excelName();
				}
				excelInfo[1] = excelDetails.sheetName();
			}
			methodExcelDetails.put(method.getName(), excelInfo);
		}
	}

	public void getExcelDetails() {
		getExcelDetailsFromClass();
		getExcelDetailsFromMethod();
	}

	public Object[][] data(Method method) {
		if (methodExcelDetails.get(method.getName())[0] != null
				&& methodExcelDetails.get(method.getName())[1] != null) {
			String[] methodExcelInfo = methodExcelDetails.get(method.getName());
			return ReadExcel.readData(methodExcelInfo);
		}
		return ReadExcel.readData(excelDetailsValue);
	}

}
