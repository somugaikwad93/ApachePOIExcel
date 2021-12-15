package com.exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelData {

	public static void main(String[] args) throws IOException {
		String path = ".\\Files\\Excel1.xlsx";
		FileInputStream fis = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);

		// using for loop

		int rows = sheet.getLastRowNum(); // gets number of rows present in the sheet
		int cols = sheet.getRow(1).getLastCellNum(); // gets number of columns present in the sheet

		// 1st for loop represents for rows and 2nd for loop represents for
		// columns/cells
		for (int r = 0; r < rows; r++) {
			XSSFRow row = sheet.getRow(r);
			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.getCell(c);
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				}
				System.out.print(" ");
			}
			System.out.println();
		}

		// using Iterator
		Iterator<Row> rows1 = sheet.iterator();
		while (rows1.hasNext()) {
			XSSFRow row = (XSSFRow) rows1.next();
			Iterator<Cell> cols1 = row.cellIterator();
			while (cols1.hasNext()) {
				XSSFCell cell = (XSSFCell) cols1.next();
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				}
				System.out.print(" ");
			}
			System.out.println();
		}
	}
}
