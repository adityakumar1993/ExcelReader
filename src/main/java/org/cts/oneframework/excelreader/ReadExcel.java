package org.cts.oneframework.excelreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
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

public class ReadExcel {

	private static XSSFSheet xssfSheet;
	private static FileInputStream fis;
	private static DataFormatter dataFormatter = new DataFormatter();
	private static Logger logger = Logger.getLogger(ReadExcel.class.getName());

	private ReadExcel() {
	}

	private static void setup(String fileName, String sheetName) throws IOException {
		
		ClassLoader loader = Thread.currentThread().getContextClassLoader();
		File folderPath = new File(loader.getResource("./data").getFile());
		File xlsFile = new File(folderPath + File.separator + fileName + ".xls");
		File xlsxFile = new File(folderPath + File.separator + fileName + ".xlsx");
		if (xlsFile.exists()) {
			fis = new FileInputStream(xlsFile);
		} else if (xlsxFile.exists()) {
			fis = new FileInputStream(xlsxFile);
		} else {
			throw new IOException("ExcelDetails annotation may be missing or excel file/sheet doesn't exists.");
		}
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
		xssfSheet = xssfWorkbook.getSheet(sheetName);
		xssfWorkbook.close();
	}

	public static Object[][] readData(String[] excelInfo) {
		String excelName = excelInfo[0];
		String sheetName = excelInfo[1];
		List<Object[]> results = new ArrayList<>();
		try {
			setup(excelName, sheetName);
			int numRows = xssfSheet.getLastRowNum();
			for (int i = 1; i <= numRows; i++) {
				Map<String, String> inputValues = getHashMapDataFromRow(xssfSheet, i);
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

	public static List<HashMap<String, String>> readData(String excelName, String sheetName) {
		List<HashMap<String, String>> excelData = new ArrayList<>();
		try {
			setup(excelName, sheetName);
			int numRows = xssfSheet.getLastRowNum();
			for (int i = 1; i <= numRows; i++) {
				HashMap<String, String> inputValues = getHashMapDataFromRow(xssfSheet, i);
				excelData.add(inputValues);
			}
		} catch (IOException e) {
			logger.warning(e.getMessage());
		} finally {
			IOUtils.closeQuietly(fis);
		}
		return excelData;
	}

	private static HashMap<String, String> getHashMapDataFromRow(Sheet sheet, int rowIndex) {
		HashMap<String, String> results = new HashMap<>();
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
