package org.data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {

		File loc = new File("F:\\Eclipse-workspace\\DataDriven\\Input.xlsx");
		FileInputStream input = new FileInputStream(loc);
		Workbook work = new XSSFWorkbook(input);

		Sheet sheet = work.getSheet("Sheet1");

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {

			Row row = sheet.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {

				Cell cell = row.getCell(j);

				int type = cell.getCellType();

				String value = null;

				if (type == 1) {

					value = cell.getStringCellValue();

				} else if (type == 0) {

					if (DateUtil.isCellDateFormatted(cell)) {

						Date d = cell.getDateCellValue();

						SimpleDateFormat f = new SimpleDateFormat("MM/DD/YYYY");
						value = f.format(d);

					}

					else {

						double d = cell.getNumericCellValue();

						long l = (long) d;

						value = String.valueOf(l);

					}

				}

				System.out.println(value);

			}

		}

	}

}
