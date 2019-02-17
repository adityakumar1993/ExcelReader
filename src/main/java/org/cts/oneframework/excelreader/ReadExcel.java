package org.cts.oneframework.excelreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.cts.oneframework.exception.ExcelDetailException;

public class ReadExcel {

	private static XSSFSheet xssfSheet;
	private static FileInputStream fis;
	private static DataFormatter dataFormatter = new DataFormatter();
	private static Logger logger = Logger.getLogger(ReadExcel.class.getName());

	private ReadExcel() {
	}

	private static void setup(String excelName, String sheetName) throws IOException {

		if (excelName == null && sheetName == null) {
			throw new ExcelDetailException("ExcelDetails annotation is missing. It must be called at either Method level or class level. If both available, method level will have the priority over class level.");

		}
		ClassLoader loader = Thread.currentThread().getContextClassLoader();
		File folderPath = new File(loader.getResource("./data").getFile());
		File xlsFile = new File(folderPath + File.separator + excelName + ".xls");
		File xlsxFile = new File(folderPath + File.separator + excelName + ".xlsx");
		if (xlsFile.exists()) {
			fis = new FileInputStream(xlsFile);
		} else if (xlsxFile.exists()) {
			fis = new FileInputStream(xlsxFile);
		} else {
			throw new ExcelDetailException("Excel Details are not correct. Trying to load excel '" + excelName + "' and sheet name '" + sheetName + "'. Either or both of which are not available/wrong.");
		}
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
		xssfSheet = xssfWorkbook.getSheet(sheetName);
		if (xssfSheet == null) {
			xssfWorkbook.close();
			throw new ExcelDetailException("Excel Sheet name is not correct. Trying to load sheet '" + sheetName + "' from excel '" + excelName + "', which looks like not available.");
		}
		xssfWorkbook.close();
	}

	public static Object[][] getData(String excelName, String sheetName) {
		List<Object[]> results = new ArrayList<>();
		try {
			setup(excelName, sheetName);
			int numRows = xssfSheet.getLastRowNum();
			for (int i = 1; i <= numRows; i++) {
				Map<String, String> inputValues = getMapDataFromRow(xssfSheet, i);
				results.add(new Object[] { inputValues });
			}
		} catch (IOException e) {
			logger.warning(e.getMessage());
		}

		finally {
			IOUtils.closeQuietly(fis);
		}
		return results.toArray(new Object[0][]);
	}

	public static List<Map<String, String>> readData(String excelName, String sheetName) {
		List<Map<String, String>> excelData = new ArrayList<>();
		try {
			setup(excelName, sheetName);
			int numRows = xssfSheet.getLastRowNum();
			for (int i = 1; i <= numRows; i++) {
				Map<String, String> inputValues = getMapDataFromRow(xssfSheet, i);
				excelData.add(inputValues);
			}
		} catch (IOException e) {
			logger.warning(e.getMessage());
		} finally {
			IOUtils.closeQuietly(fis);
		}
		return excelData;
	}

	private static Map<String, String> getMapDataFromRow(Sheet sheet, int rowIndex) {
		Map<String, String> results = new LinkedHashMap<>();
		String[] columnHeaders = getDataFromRow(sheet, 0);
		String[] valuesFromRow = getDataFromRow(sheet, rowIndex);
		for (int i = 0; i < columnHeaders.length; i++) {
			if (i >= valuesFromRow.length) {
				results.put(columnHeaders[i], "");
			} else {
				results.put(columnHeaders[i], valuesFromRow[i]);
			}
		}
		return results;
	}

	private static String[] getDataFromRow(Sheet sheet, int rowIndex) {
		FormulaEvaluator formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		Row row = sheet.getRow(rowIndex);
		short numCells = row.getLastCellNum();
		String[] result = new String[numCells];
		for (int i = 0; i < numCells; i++) {
			result[i] = getValueAsString(row.getCell(i), formulaEvaluator);
		}
		return result;
	}

	private static String getValueAsString(Cell cell, FormulaEvaluator formulaEvaluator) {
		if (cell != null) {
			CellType cellType = cell.getCellTypeEnum();
			if (cellType.equals(CellType.BOOLEAN)) {
				return String.valueOf(cell.getBooleanCellValue());
			} else if (cellType.equals(CellType.NUMERIC)) {
				return dataFormatter.formatCellValue(cell);
			} else if (cellType.equals(CellType.STRING)) {
				return cell.getRichStringCellValue().getString();
			} else if (cellType.equals(CellType.FORMULA)) {
				return formulaEvaluator.evaluate(cell).getStringValue();
			}
		}
		return "";
	}

}
