package com.excelfiles.ExcelFileReadAndWrite;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {
	public static void main(String[] args) throws Exception {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Details");
		sheet.createRow(0);
		sheet.getRow(0).createCell(0).setCellValue("Emp_Id");
		sheet.getRow(0).createCell(1).setCellValue("Name");

		sheet.createRow(1);
		sheet.getRow(1).createCell(0).setCellValue("567");
		sheet.getRow(1).createCell(1).setCellValue("ABC");

		File file = new File("F:\\Java Practice_Selenium\\ExcelFileReadAndWrite\\ExcelFiles\\Emp_Details.xlsx");
		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);

		workbook.close();

	}
}
