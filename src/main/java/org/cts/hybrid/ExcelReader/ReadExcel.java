package org.cts.hybrid.ExcelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	private static XSSFWorkbook xssfWorkbook;
	private static XSSFSheet xssfSheet;
	private static FileInputStream fis;
	private static DataFormatter dataFormatter = new DataFormatter();
	private static FormulaEvaluator formulaEvaluator;

	private static void setup(String fileName, String sheetName) throws IOException {
		ClassLoader loader = Thread.currentThread().getContextClassLoader();
		File folderPath = new File(loader.getResource("./data").getFile());
		File xlsFile = new File(folderPath + "/" + fileName + ".xls");
		File xlsxFile = new File(folderPath + "/" + fileName + ".xlsx");
		if (xlsFile.exists()) {
			fis = new FileInputStream(xlsFile);
		} else if (xlsxFile.exists()) {
			fis = new FileInputStream(xlsxFile);
		} else {
			throw new FileNotFoundException("File doesn't exists:");
		}
		xssfWorkbook = new XSSFWorkbook(fis);
		xssfSheet = xssfWorkbook.getSheet(sheetName);
	}

	static Object[][] readData(String fileName, String sheetName) {
		List<Object[]> results = new ArrayList<Object[]>();
		try {
			setup(fileName, sheetName);
			int numRows = xssfSheet.getLastRowNum();
			for (int i = 1; i <= numRows; i++) {
				Map<String, String> inputValues = getHashMapDataFromRow(xssfSheet, i);
				if (inputValues == null) {
					break;
				}
				results.add(new Object[] { inputValues });
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

		finally {
			IOUtils.closeQuietly(fis);
		}
		return results.toArray(new Object[0][]);
	}

	private static Map<String, String> getHashMapDataFromRow(Sheet sheet, int rowIndex) {
		Map<String, String> results = new HashMap<String, String>();
		String[] columnHeaders = getDataFromRow(sheet, 0);
		String[] valuesFromRow = getDataFromRow(sheet, rowIndex);
		if (valuesFromRow == null) {
			return null;
		}
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
		formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			return null;
		}
		short numCells = row.getLastCellNum();
		String[] result = new String[numCells];
		for (int i = 0; i < numCells; i++) {
			result[i] = getValueAsString(row.getCell(i), formulaEvaluator);
		}
		return result;
	}

	private static String getValueAsString(Cell cell, FormulaEvaluator formulaEvaluator) {
		CellType cellType = CellType._NONE;

		if (cell != null) {
			cellType = cell.getCellTypeEnum();
		}

		if (cellType.equals(CellType.BLANK)) {
			return "";
		} else if (cellType.equals(CellType.ERROR)) {
			return "";
		} else if (cellType.equals(CellType.BOOLEAN)) {
			return String.valueOf(cell.getBooleanCellValue());
		} else if (cellType.equals(CellType.NUMERIC)) {
			return dataFormatter.formatCellValue(cell);
		} else if (cellType.equals(CellType.STRING)) {
			return cell.getRichStringCellValue().getString();
		} else if (cellType.equals(CellType.FORMULA)) {
			String value = formulaEvaluator.evaluate(cell).getStringValue();
			return value;
		} else {
			return "";
		}

	}

}
