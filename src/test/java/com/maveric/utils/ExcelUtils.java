package com.maveric.utils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static XSSFSheet sh;
	public static void main(String[] args) throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		XSSFWorkbook wb= new XSSFWorkbook();
		//XSSFSheet sh=wb.createSheet("Data");
		sh = wb.createSheet("Data");
		//XSSFRow row=sh.createRow(0);
		//XSSFCell cell=row.createCell(0);
		//index value is Zero
		//cell.setCellValue("123");
		//cell=row.createCell(1);
		//cell.setCellValue("123");
		//row=sh.createRow(1);
		//cell=row.createCell(0);
		//cell.setCellValue(true);
		addData(1,1,"123");
		addData(1,2,1234);
		addData(2,1,true);
		addData(2,2,12.12);
		wb.write(new FileOutputStream("src/test/resources/Data/TestData.xlsx"));
		wb.close();
		
	}
	
	public static void setExcel(String workbook)
	
	public static void addData(int row_count, int col_count, Object data)
	{
		XSSFRow row = null;
		if(sh.getRow(row_count-1)==null)
		{
			row = sh.createRow(row_count-1);
		}
	}

}
