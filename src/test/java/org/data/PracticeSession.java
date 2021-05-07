package org.data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PracticeSession {

	public static void main(String[] args) throws IOException {

		File file = new File("F:\\Eclipse-workspace\\DataDriven\\Practice.xlsx");
		FileInputStream stream = new FileInputStream(file);

		Workbook work = new XSSFWorkbook(stream);

		/*
		 * Sheet sheet = work.getSheet("Sheet1");
		 * 
		 * Row row = sheet.getRow(0);
		 * 
		 * Cell cell = row.getCell(0);
		 * 
		 * cell.setCellValue("jackmavles");
		 * 
		 * FileOutputStream out = new FileOutputStream(file);
		 * 
		 * work.write(out);
		 * 
		 * System.out.println("done");
		 */
		Sheet sheet = work.createSheet("sheet2");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		cell.setCellValue("selvam");

		FileOutputStream out = new FileOutputStream(file);

		work.write(out);
		
		System.out.println("done");

	}

}
