package com.exceloperations;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;

public class WritingExcelData {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info"); //creates sheet
		
		//datas to add on the sheet
		Object empdata[][] = {{"EmpId","Name","Attendance"},
							{101 , "David", true},
							{102,"John" , true},
							{103,"Smith", false}
							};
		
		int rows = empdata.length; //No of rows
		int cols = empdata[0].length; //No of columns
		
		for(int r=0 ; r<rows ; r++) 
		{
			XSSFRow row = sheet.createRow(r); //creates Rows
			for(int c=0 ; c<cols ; c++) 
			{
				XSSFCell cell = row.createCell(c); //creates columns
				Object value = empdata[r][c];
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
		
		String path = "./Files/emp.xlsx";
		FileOutputStream fos = new FileOutputStream(path);
		workbook.write(fos);
		fos.close();

	}

}
