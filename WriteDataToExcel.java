package com.xworkz.reading.writeopr;

import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteDataToExcel {

	// any exceptions need to be caught
	public static void main(String[] args) throws Exception {
		// workbook object
		XSSFWorkbook workbook = new XSSFWorkbook();

		// spreadsheet object
		XSSFSheet spreadsheet = workbook.createSheet(" Sports Data ");

		// creating a row object
		XSSFRow row;

		// This data needs to be written (Object[])
		Map<String, Object[]> sportsData = new TreeMap<String, Object[]>();

		sportsData.put("1", new Object[] { "ID", "SPORTS", "PLAYER", "COUNTRY", "AGE" });

		sportsData.put("2", new Object[] { "1", "Cricketer", "MSdhoni", "India", "40" });

		sportsData.put("3", new Object[] { "2", "Tennis", "Roger Federer", "Switzerland", "40" });

		sportsData.put("4", new Object[] { "3", "Badminton", "P.V.Sindhu", "India", "26" });

		sportsData.put("5", new Object[] { "4", "Hockey", "Harmanpreet Singh", "India", "26" });

		sportsData.put("6", new Object[] { "5", "Cricketer", "KLRahul", "India", "29" });

		Set<String> keyid = sportsData.keySet();

		int rowid = 0;

		// writing the data into the sheets...

		for (String key : keyid) {

			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = sportsData.get(key);
			int cellid = 0;

			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}

		// .xlsx is the format for Excel Sheets...
		// writing the workbook into the file...
		FileOutputStream out = new FileOutputStream(new File("D:\\EXCELjava\\sport.xlsx"));

		workbook.write(out);
		out.close();
	}
}
