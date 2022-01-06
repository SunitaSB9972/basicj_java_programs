package com.xworkz.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;


public class ReadExcelFileDemo {
	
  public static void main(String[] args)throws IOException {
	  
	        InputStream ExcelFileToRead = new FileInputStream("D:\\EXCELjava\\sports.xlsx");
	        XSSFWorkbook  wb = new XSSFWorkbook(ExcelFileToRead);

	        XSSFWorkbook test = new XSSFWorkbook(); 

	        XSSFSheet sheet = wb.getSheetAt(0);
	        XSSFRow row; 
	        XSSFCell cell;

	        Iterator rows = sheet.rowIterator();

	        while (rows.hasNext())
	        {
	            row=(XSSFRow) rows.next();
	            Iterator cells = row.cellIterator();
	            while (cells.hasNext())
	            {
	                cell=(XSSFCell) cells.next();

	                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
	                {
	                    System.out.print(cell.getStringCellValue()+" ");
	                }
	                else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
	                {
	                    System.out.print(cell.getNumericCellValue()+" ");
	                }
	                else
	                {
	                    //U Can Handel Boolean, Formula, Errors
	                }
	            }
	            System.out.println();
	        }

	    }
}