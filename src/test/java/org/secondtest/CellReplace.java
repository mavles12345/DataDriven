package org.secondtest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellReplace {

	public static void main(String[] args) throws IOException {

		File f = new File("F:\\Eclipse-workspace\\DataDriven\\Seconddatafile.xlsx");

		FileInputStream input = new FileInputStream(f);

		Workbook work = new XSSFWorkbook(input);

		Sheet sheet = work.getSheet("Sheet1");

		Row row = sheet.getRow(0);

		Cell cell = row.getCell(0);

		System.out.println(cell);

		cell.setCellValue("Mavles");

		FileOutputStream out = new FileOutputStream(f);

		work.write(out);
		
		System.out.println("done");

	}

}
