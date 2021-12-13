package com.excelfiles.ExcelFileReadAndWrite;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {
	public static void main(String[] args) throws Exception {
		File file = new File("F:\\Java Practice_Selenium\\ExcelFileReadAndWrite\\ExcelFiles\\Emp_Details.xlsx");
		FileInputStream outputStream = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(outputStream);
		XSSFSheet sheet = wb.getSheet("Details");

//		double i =sheet.getRow(1).getCell(0).getNumericCellValue();
//		System.out.println(i);

		int rowCount = sheet.getPhysicalNumberOfRows();

		for (int i = 0; i < rowCount; i++) {
			XSSFRow row = sheet.getRow(i);

			int colCount = row.getPhysicalNumberOfCells();
			for (int j = 0; j < colCount; j++) {
				XSSFCell cell = row.getCell(j);
				String cellValue = getCellValue(cell);
				System.out.print("||"+cellValue);
			}
			System.out.println();
		}
	}

	public static String getCellValue(XSSFCell cell) {

		switch (cell.getCellType()) {
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());

		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());

		case STRING:
			return cell.getStringCellValue();

		default:
			return cell.getStringCellValue();
		}
	}

}
