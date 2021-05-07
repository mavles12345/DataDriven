package org.data;

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

public class CreateSheet {

	public static void main(String[] args) throws IOException {
		
		//File loc=new File("F:\Eclipse-workspace\DataDriven\Update.xlsx");
		File loc = new File("F:\\Eclipse-workspace\\DataDriven\\Update.xlsx");
		
	//	FileInputStream input=new FileInputStream(loc);
		
		Workbook work=new XSSFWorkbook();
		
		Sheet sheet = work.createSheet("Sheet3");
		
		Row row=sheet.createRow(0);
		
		Cell cell=row.createCell(0);
		
		cell.setCellValue("VijiThishiHenik");
		
		FileOutputStream out=new FileOutputStream(loc);
		
		work.write(out);
		
		
		
		
	}
	
}
